# Security policy

## Supported versions

Only the latest published release on npm receives security fixes. The project is pre-1.0, so there are no maintained release lines.

## Reporting a vulnerability

Please do not open a public issue for security problems.

Use GitHub's private vulnerability reporting: open the **Security** tab of the repository and choose **Report a vulnerability**. If that is not available to you, email randeep@lucioai.com with "docxtor security" in the subject.

Include what you can of: the affected version, a description of the issue, and a minimal document or fixture code that reproduces it. Please do not send real documents that contain confidential content.

You will get an acknowledgement within five working days. We aim to ship a fix for confirmed issues within thirty days and will credit you in the release notes unless you ask otherwise.

## Scope and threat model

docxtor parses `.docx` files, which are ZIP archives of XML parts. Things worth knowing when you process untrusted input:

- The whole archive is read into memory with JSZip, and every XML part docxtor touches is parsed into JavaScript objects. Apply size limits before handing files to the library if callers are untrusted.
- Patch targets and text come from whatever produced them, typically a language model. The validator rejects malformed patches, but it does not sanitise content: text is written into the document as given, with XML entities escaped.
- docxtor never executes macros, resolves external relationships, or fetches anything over the network.

Reports about denial of service through crafted archives or XML are in scope. Reports about the content a language model chose to write into a document are not.
