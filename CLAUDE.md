# CLAUDE.md - Project Instructions

## Project
SuperQAT — an Office Add-in (Web + COM/VBA) that exposes all 2,199 Office 365 commands with search and categories for Word, Excel, and PowerPoint.

## Communication Style
- Explain things simply and clearly, as if the user is five years old.
- Give exact commands to copy-paste — don't assume the user knows which directory they're in or what flags mean.
- When giving Termux commands, always start with `cd ~/qat-exposer` to make sure we're in the right place.

## Git Workflow
- Always merge your feature branch into `main` before finishing.
- After all work is done, push `main` to the remote.
- Do NOT leave changes only on a feature branch — `main` must always have the latest code.
- Steps: commit on feature branch → merge feature branch into main → push main → push feature branch.

## README.md
- Update README.md every time a release is prepared or significant changes are made.
- Keep the command count, platform support table, project structure, and download links accurate.
- If the version number changes, update it in README.md too.
