/**
 * Content controls (`w:sdt`): reading a control's properties into a
 * ControlInfo, and resolving the chain of wrappers a walker crossed into the
 * innermost control.
 */

import type { FXPNode, ControlInfo } from "../types.js";
import { findChild, getAttr, getChildren, getTagName } from "../xml/helpers.js";

/** `w:sdtPr` children that say what kind of control this is. */
const KIND_BY_TAG: Record<string, ControlInfo["kind"]> = {
  "w:richText": "richText",
  "w:text": "text",
  "w:date": "date",
  "w:dropDownList": "dropDownList",
  "w:comboBox": "comboBox",
  "w14:checkbox": "checkbox",
  "w:picture": "picture",
  "w:docPartObj": "docPart",
  "w:docPartList": "docPart",
  "w:equation": "other",
  "w:citation": "other",
  "w:group": "other",
  "w:bibliography": "other",
  "w15:repeatingSection": "other",
  "w15:repeatingSectionItem": "other",
};

/** True for `<w:x/>` and `<w:x w:val="1|true|on"/>`. */
function isOn(node: FXPNode, attr: string): boolean {
  const val = getAttr(node, attr);
  return val === undefined || (val !== "0" && val !== "false" && val !== "off");
}

/**
 * Read a `w:sdt` node's properties.
 *
 * A control with no type element in `w:sdtPr` is a rich-text control: that
 * is what Word writes for one.
 */
export function readControlInfo(sdtNode: FXPNode, parent?: ControlInfo): ControlInfo {
  const info: ControlInfo = { id: "", kind: "richText", placeholder: false, xmlRef: sdtNode };
  if (parent) info.parent = parent;

  const sdtPr = findChild(sdtNode, "w:sdtPr");
  if (!sdtPr) return info;

  for (const child of getChildren(sdtPr)) {
    const tag = getTagName(child);
    if (!tag) continue;

    if (tag === "w:id") {
      info.id = getAttr(child, "w:val") ?? "";
    } else if (tag === "w:alias") {
      const val = getAttr(child, "w:val");
      if (val !== undefined) info.alias = val;
    } else if (tag === "w:tag") {
      const val = getAttr(child, "w:val");
      if (val !== undefined) info.tag = val;
    } else if (tag === "w:showingPlcHdr") {
      info.placeholder = isOn(child, "w:val");
    } else if (tag === "w:lock") {
      const val = getAttr(child, "w:val");
      if (val === "contentLocked") info.locked = "content";
      else if (val === "sdtLocked") info.locked = "control";
      else if (val === "sdtContentLocked") info.locked = "both";
    } else if (tag === "w:dataBinding" || tag === "w15:dataBinding") {
      info.bound = true;
    } else if (tag in KIND_BY_TAG) {
      info.kind = KIND_BY_TAG[tag]!;
      if (tag === "w:dropDownList" || tag === "w:comboBox") {
        info.options = getChildren(child)
          .filter((c) => getTagName(c) === "w:listItem")
          .map((c) => getAttr(c, "w:displayText") ?? getAttr(c, "w:value") ?? "");
      } else if (tag === "w14:checkbox") {
        const checked = findChild(child, "w14:checked");
        info.checked = checked ? isOn(checked, "w14:val") : false;
      }
    }
  }

  return info;
}

/**
 * Turns the `w:sdt` wrappers a walker crossed into ControlInfo objects, one
 * per `w:sdt` node, so every paragraph and cell inside the same control
 * shares the same object.
 */
export class ControlResolver {
  private readonly seen = new Map<FXPNode, ControlInfo>();

  /**
   * The innermost control of a chain of wrappers (outermost first) that sits
   * inside `outer`. Returns `outer` for an empty chain.
   */
  resolve(sdts: FXPNode[], outer?: ControlInfo): ControlInfo | undefined {
    let current = outer;
    for (const sdt of sdts) {
      let info = this.seen.get(sdt);
      if (!info) {
        info = readControlInfo(sdt, current);
        this.seen.set(sdt, info);
      }
      current = info;
    }
    return current;
  }
}

/** A control and the controls that hold it, innermost first. */
export function controlChain(control: ControlInfo | undefined): ControlInfo[] {
  const chain: ControlInfo[] = [];
  for (let c = control; c; c = c.parent) chain.push(c);
  return chain;
}
