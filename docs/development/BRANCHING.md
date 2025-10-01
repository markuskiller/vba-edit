# BRANCHING.md

## 🌿 Branching Strategy Overview

This document outlines the branching strategy for my fork of the upstream repository. It helps me keep track of:

- Which changes are intended for upstream contribution via Pull Requests
- Which changes are personal and should remain in my fork
- How my branches relate to the upstream `main` and `dev` branches

---

## 🔹 Primary Branches

| Branch         | Purpose                                                                 |
|----------------|-------------------------------------------------------------------------|
| `main`         | My personal development base. Contains all my custom changes.           |
| `upstream-dev` | Tracks the latest state of the upstream repository's `dev` branch.      |

> I regularly update `upstream-dev` using:
> `git fetch upstream && git checkout upstream-dev && git pull`

---

## 🔹 Feature Branches

Feature branches are created either from `upstream-dev` (for upstream contributions) or from `main` (for personal enhancements).

| Branch Name              | Based On      | Purpose                                           | Intended for PR? |
|--------------------------|---------------|---------------------------------------------------|------------------|
| `feature-docs-newbies`   | `main`        | Additional documentation for new developers       | ❌               |
| `feature-program-option` | `upstream-dev`| New program option requested in upstream issues   | ✅               |
| `feature-dev-config`     | `main`        | Development configuration changes                 | ❌               |
| `feature-refactor-clean` | `upstream-dev`| Structural refactoring to reduce code redundancy  | ✅ (if isolated) |

---

## 🔧 Workflow Summary

1. **Start a new feature**
   - Decide if it's for upstream or personal use
   - Create a branch from `upstream-dev` or `main` accordingly

2. **Develop and commit**
   - Keep commits focused and meaningful
   - Use `git rebase -i` to clean up history if needed

3. **Push and open Pull Request**
   - Push to my fork: `git push origin feature-xyz`
   - Open a PR to `upstream/dev` if applicable

4. **Stay in sync with upstream**
   - Regularly update `upstream-dev`
   - Rebase feature branches onto `upstream-dev` to stay current

---

## 🧩 Notes

- I use `main` as my sandbox for personal development.
- I split features into separate branches to keep things modular and reviewable.
- I document each feature branch with its purpose and status (local, pushed, PR open, merged).
