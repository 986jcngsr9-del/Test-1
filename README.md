# Test-1 — Learning Claude Code

This is a test repository I'm using to teach myself how to use **Claude Code on the web**.

## What this is

I'm chatting with Claude Code inside an isolated cloud container that cloned this
repo. Claude can read and write code, run commands, and commit & push to a branch.
The container is ephemeral, so anything worth keeping has to be committed and
pushed — which Claude handles. **When we push, it goes to this GitHub repo.**

## How to start using it

Just tell Claude what you want in plain language. For example:

- **"Create a starter Node/Python/etc. project"** — it scaffolds files, commits, and pushes.
- **"Add a README describing X"** — it writes and commits it.
- **"Build a small app that does Y"** — it plans, implements, and pushes.
- **"Set up tests and a linter"** — it can wire those up, and even add a SessionStart
  hook so they run automatically in future web sessions.

## Pull requests & CI

Once there are changes on a branch, I can ask Claude to **open a pull request**
(it won't create one unless asked). Claude can also **watch the PR** to respond to
review comments and fix failing CI automatically.

## The basic loop

1. Ask Claude for what I want, in plain language.
2. Claude makes the changes on a branch.
3. Claude commits and pushes to GitHub.
4. (Optional) Open a PR and merge.

## Notes to self

- The working branch for this session is `claude/getting-started-okpvk3`.
- Nothing is saved unless it's committed and pushed.
- I can ask "what can you do?" any time to explore more.
