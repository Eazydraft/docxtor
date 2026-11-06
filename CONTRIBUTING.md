# Contributing to docxtor

Thanks for your interest. This guide covers how to set up the project, what a good change looks like, and the sign-off every commit needs.

## Getting started

Requirements: Node.js 20.19 or newer and [pnpm](https://pnpm.io) 10.

```bash
git clone https://github.com/Eazydraft/docxtor.git
cd docxtor
pnpm install
pnpm test
```

Useful scripts:

| Script | What it does |
|---|---|
| `pnpm test` | Run the test suite with Vitest |
| `pnpm test:watch` | Re-run tests on change |
| `pnpm typecheck` | Type-check with `tsc --noEmit` |
| `pnpm build` | Build ESM, CJS and type declarations with tsdown |
| `pnpm lint` | Check the published package shape with publint and arethetypeswrong |
| `pnpm check` | Everything CI runs: typecheck, test, build, lint |

## Making a change

1. Open an issue first for anything larger than a bug fix, so the approach can be agreed before you write code.
2. Branch from `main`.
3. Keep the change focused. One concern per pull request.
4. Add or update tests. Test fixtures are built in code in `tests/fixtures/create-docx.ts`; do not commit `.docx` files.
5. Update the relevant page under `docs/`.
6. Run `pnpm check` locally.

### Invariants to keep

- The normalized AST is a read-only view. Patches mutate the underlying XML objects through `xmlRef` pointers, then the affected paragraph is re-normalized. Never write back from the AST.
- Content a patch does not touch must survive unchanged: styles, numbering, section properties, images, unknown elements.
- Positional ids (`p0`, `tbl0.r1.c0.p0`) are stable within a session. Native `w14:paraId` values are never rewritten.
- With `trackRevisions`, applying a batch and then accepting all revisions must give the same result as applying the batch untracked. Tests pin `revisionDate` to compare XML byte for byte.
- Atomic batches must leave the document untouched if any patch fails.

### Commit messages

Use the [Conventional Commits](https://www.conventionalcommits.org) prefixes already in the history: `feat:`, `fix:`, `docs:`, `test:`, `chore:`, `ci:`, `refactor:`. The first line should say what changed, not how.

## Developer Certificate of Origin

This project uses the [Developer Certificate of Origin](DCO) (DCO) instead of a contributor license agreement. By signing off a commit you certify that you wrote the change or otherwise have the right to submit it under the MIT license.

Every commit must carry a `Signed-off-by` line with your real name and email:

```
Signed-off-by: Jane Doe <jane@example.com>
```

Git adds it for you with the `-s` flag:

```bash
git commit -s -m "fix: keep page breaks on the original side of a rebuilt run"
```

To make it automatic, add a `prepare-commit-msg` hook or set a git alias such as `git config alias.ci 'commit -s'`.

If you forget, amend the last commit with `git commit --amend -s --no-edit`, or sign off a whole branch with `git rebase --signoff main`. The DCO check on pull requests will not pass until every commit is signed.

## Pull requests

- Fill in the pull request template.
- CI must pass: typecheck, tests on all supported Node versions, build and package lint.
- A maintainer will review. Expect requests for tests or documentation if they are missing.
- Squash merges are the norm, so the PR title should itself be a good commit message.

## Reporting bugs

Use the bug report template. A failing test or a snippet that builds the fixture in code is the fastest path to a fix. Please do not attach real documents that contain confidential content.

For security issues, follow [SECURITY.md](SECURITY.md) instead of opening a public issue.

## Releases (maintainers)

1. Bump `version` in `package.json`.
2. Merge, then create a GitHub release with tag `vX.Y.Z` matching the package version.
3. The publish workflow verifies the tag, runs the checks and publishes to npm with provenance.
