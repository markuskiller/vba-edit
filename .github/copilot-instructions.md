# Project Overview

This project is a command line application that allows users to export, import and synchronize custom program code (VBA and PowerQuery) contained in Microsoft Office files.

## Folder Structure

- `/src`: Contains the source code
- `/docs`: Contains documentation for the project, including ideas for further enhancements and even programming guides for beginners.

## General Instructions

The project's language is English.

Write code (names of classes, objects, variables, methods, comments etc) always in the project language.

Also express all your answers in project language, even when prompted in any other language.
As an exception, the developer may ask for a translation of terms and sentences he or she doesn't understand, on his/her behalf.

## Version Control

Currently, some developers are new to working with Git, Github and Copilot.
When working on a feature, suggest which branch to use.

When the developer asks questions about tools he uses, switch to a different chat reserved for these questions.

Microsoft Copilot suggested the branching strategy outlined in .\docs\development\BRANCHING.md

## Github Copilot Modes

When in Agent mode, don't modify any code, unless you did check out to a new Git branch before.
Each prompt shall use the context @workspace automatically.

## CHANGELOG Guidelines

When updating CHANGELOG.md:

### Read First
- **Always read the entire CHANGELOG thoroughly** before making changes
- Check for duplicates - don't repeat existing entries
- Understand the context of what's already documented

### Writing Style
- **User-focused**: Write for end users, not developers
- **Benefits-oriented**: Explain what it means for users, not how it works internally
- **Simple language**: Avoid technical jargon and implementation details
- **Concise**: Each entry should be clear and brief

### What to Include
- ✅ New features users can see and use
- ✅ Changes to existing behavior users will notice
- ✅ Bug fixes that affected users
- ✅ Security improvements that matter to users
- ✅ Deprecated features with migration guidance
- ✅ Breaking changes (clearly marked)

### What to Exclude
- ❌ Internal refactoring (unless it impacts performance/behavior)
- ❌ Test changes or additions
- ❌ Code architecture changes
- ❌ Function/class names and technical implementation
- ❌ Developer-only improvements
- ❌ CI/CD changes (unless they affect releases/binaries)

### Format Guidelines
- Use **bold** for feature names
- Keep entries action-oriented ("Added", "Fixed", "Changed")
- Include issue numbers with links: `[Issue #20](https://github.com/...)`
- Credit contributors: `([@username](https://github.com/username))`
- Group related changes together

### Categories
- **Added**: New features
- **Changed**: Modifications to existing features
- **Deprecated**: Features being phased out
- **Removed**: Features that were removed
- **Fixed**: Bug fixes
- **Security**: Security-related improvements

### Examples

❌ **Too Technical**:
> "Refactored CLI architecture with inline argument definitions using EnhancedHelpFormatter class"

✅ **User-Focused**:
> "Improved help messages with organized option groups and practical examples"

❌ **Internal Detail**:
> "Added `add_folder_organization_arguments()` function in `cli_common.py`"

✅ **User Benefit**:
> "Folder options now appear only on commands where they're relevant"

### Special Notes
- Security improvements get their own **Security** section
- Always check if Windows binary/release features should be mentioned
- Dependabot and automation features belong in **Added** if user-facing
- Breaking changes should be very clearly called out

## Documentation Strategy

### Current Approach (Alpha/Beta Stage)
- **Minimal public documentation**: Focus on README and excellent CLI help messages
- **Keep development docs private**: Testing guides, build processes, and internal documentation remain in local/gitignored `docs/` folder
- **Resource-conscious**: Limited documentation resources during alpha/beta development
- **Quality over quantity**: Better to have great README and help than mediocre comprehensive docs

### Public Documentation (What to Include)
- ✅ **README.md**: Primary user-facing documentation
- ✅ **CLI help messages**: Well-organized, grouped options with examples
- ✅ **CHANGELOG.md**: User-focused release notes
- ✅ **SECURITY.md**: Security policy and vulnerability reporting
- ✅ **LICENSE**: Project license and third-party licenses
- ✅ **AUTHORS.md**: Contributor recognition

### Private/Local Documentation (Not for GitHub)
- 🔒 **Testing guides**: Internal test strategies and procedures
- 🔒 **Build documentation**: Binary creation, development setup
- 🔒 **Developer workflows**: Team-specific processes
- 🔒 **Architecture notes**: Internal design decisions
- 🔒 **Experimental features**: Ideas and prototypes

**Rationale**: Small team with limited resources should focus on end-user experience rather than comprehensive documentation. Internal docs help the team but don't need to be public during alpha/beta.

### Future Plans (Post-1.0)
- 📚 **ReadTheDocs**: Full documentation site when project matures
- 📚 **API documentation**: If/when programmatic usage becomes common
- 📚 **Contributing guide**: When ready for external contributors
- 📚 **Tutorials**: Once core features are stable

**Decision**: Wait until project is mature enough to warrant comprehensive public documentation.
