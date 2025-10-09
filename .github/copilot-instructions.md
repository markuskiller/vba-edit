# Project Overview

This project is a command-line application that allows users to export, import, and edit VBA code from Microsoft Office files using external editors with one-way sync back to Office.

## Current Features (v0.4.x)

**Status**: Active refinement across multiple 0.4.x releases

**Supported Applications:**
- ✅ Microsoft Excel (VBA: modules, classes, forms)
- ✅ Microsoft Word (VBA: modules, classes, forms)
- ✅ Microsoft Access (VBA: modules, classes only - NO forms)
- ✅ Microsoft PowerPoint (VBA: modules, classes, forms)

**Core Functionality:**
- **Export**: One-time export of VBA code to `.bas`, `.cls`, `.frm` files
- **Import**: One-time import of VBA code from files back to Office
- **Edit**: Live editing with **one-way sync** (external editor → Office VBA)
  - Initial export on command start
  - Watches filesystem for changes
  - Syncs changes back to Office VBA on save [CTRL+S]
  - **NO bi-directional sync**: Changes made in Office VBA editor while edit mode is active are NOT synced back to files

**Advanced Features (Being Refined):**
- RubberduckVBA folder structure support (refining)
- Configuration file support with TOML (optimizing)
- Colorized CLI output (implementing in v0.4.1)
- Safety features and data loss prevention (refining)
- Enhanced CLI with organized help and grouped options (refining)
- Windows binaries with security verification (ongoing)

## Planned Features (Future)

**Not Yet Implemented:**
- ❌ PowerQuery support (M language) - v0.5.0 (next release)
- ❌ Bi-directional sync (Office VBA → external editor) - v0.5.0 via keyboard shortcut approach
- ❌ Code signing for Windows binaries (SignPath.io) - v0.6.0+ if feasible

**Detailed Roadmap**: See `docs/ROADMAP.md` (private, not published to GitHub during alpha/beta)

**Note**: No specific date commitments - features released when ready

## Folder Structure

- `/src`: Contains the source code
- `/docs`: Contains documentation for the project, including ideas for further enhancements and even programming guides for beginners.

## General Instructions

The project's language is English.

Write code (names of classes, objects, variables, methods, comments etc) always in the project language.

Also express all your answers in project language, even when prompted in any other language.
As an exception, the developer may ask for a translation of terms and sentences he or she doesn't understand, on his/her behalf.

---

## Development Environment Setup

### CRITICAL: Always Set Up Environment First

**Before running any Python commands, tests, or scripts:**

1. **Sync dependencies** (ensures all packages including dev dependencies are installed):
   ```powershell
   uv sync --extra dev
   ```

2. **For Python commands** - Use one of these methods:
   ```powershell
   # Method 1: Use uv run (recommended - handles environment automatically)
   uv run python script.py
   uv run pytest tests/
   uv run ruff check
   
   # Method 2: Activate virtual environment manually
   .venv\Scripts\Activate.ps1  # PowerShell
   python script.py
   pytest tests/
   ```

3. **Common commands**:
   ```powershell
   # Install/update dependencies
   uv sync --extra dev
   
   # Run tests
   uv run pytest
   uv run pytest tests/test_cli_common.py -v
   
   # Code quality
   uv run ruff check --fix
   uv run ruff format
   
   # Run CLI tools
   uv run python -m vba_edit.word_vba --help
   ```

### Why This Matters

- **Missing dependencies**: Without `uv sync`, pytest, ruff, and other dev tools won't be available
- **Import errors**: Python modules won't be found if environment isn't activated
- **CI/CD parity**: Using `uv run` or activated venv matches CI/CD behavior
- **Rich dependency**: Colorization features require `rich>=13.0.0` to be installed

### Troubleshooting

If you see errors like:
- `ModuleNotFoundError: No module named 'pytest'` → Run `uv sync --extra dev`
- `ModuleNotFoundError: No module named 'rich'` → Run `uv sync`
- `Command 'pytest' not found` → Use `uv run pytest` or activate venv first

---

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

---

## Code Quality & Testing

### Pre-Test Workflow (CRITICAL)
- **ALWAYS run ruff before testing**: After adding new tests, files, or major refactoring:
  ```bash
  ruff check --fix
  ruff format
  pytest
  ```
- This prevents having to run tests multiple times due to formatting issues
- Ruff should be automatic before any test run

### Testing Practices
- All new features require corresponding tests
- Use `pytest.mark.parametrize` for testing multiple scenarios efficiently
- Use `pytest.mark.skip` for tests requiring live Office interaction (with clear reasons)
- Test files follow naming convention: `test_<module_name>.py`
- Aim for comprehensive coverage of edge cases and error conditions

### Error Handling
- Custom exception hierarchy exists in `exceptions.py` - use appropriate exceptions:
  - `OfficeError` base for Office application issues
  - `VBAError` base for VBA-specific issues
  - `VBAExportWarning` for warnings requiring user decision (doesn't inherit from VBAError)
  - Specific: `VBAAccessError`, `VBAImportError`, `VBAExportError`, `DocumentClosedError`, etc.
- Always provide clear, user-friendly error messages
- Include context about what failed and potential solutions

### Code Comments
- Use `NOTE:` for implementation notes worth highlighting
- Use `IMPORTANT:` for critical information about behavior or requirements
- Document complex logic, especially around VBA/COM interactions
- Include examples in docstrings where helpful

---

## Release Management

### Version Numbering
- Follow [Semantic Versioning](https://semver.org/)
- Alpha releases: `0.4.1a1`, `0.4.1a2`, etc.
- Beta releases: `0.4.1b1`, `0.4.1b2`, etc.
- Release candidates: `0.4.1rc1`
- Stable releases: `0.4.1`

### Release Process
- Pre-releases go through `dev` branch
- Stable releases merge `dev` → `main`
- Always update `pyproject.toml` version before release
- Update CHANGELOG.md with release date
- Use GitHub releases with appropriate templates (see `docs/development/RELEASE_PROCESS.md`)
- Binaries are built automatically on release creation

### Release Branches
- `dev`: Active development, pre-releases
- `main`: Stable releases only
- Feature branches: `feature/<name>` (merge to dev via PR)

---

## Dependencies & Tools

### Dependency Management
- Use `uv` for fast dependency installation and management
- Keep `requirements.txt` minimal (VS Code needs only)
- Main dependencies in `pyproject.toml`
- Optional dependencies (like `xlwings`) properly declared in `[project.optional-dependencies]`
- Dependabot monitors weekly for security updates

### Python Support
- Support Python 3.9 - 3.13
- Test against all supported versions in CI
- Don't use features from Python > 3.9 without compatibility checks
- **Python 3.14**: Not yet supported - waiting for `watchfiles` to release pre-built wheels for cp314

---

## Windows Binary Distribution

### PyInstaller Binaries
- Four separate executables: `excel-vba.exe`, `word-vba.exe`, `access-vba.exe`, `powerpoint-vba.exe`
- Built via `create_binaries.py` script
- Currently unsigned (expect Windows SmartScreen warnings)
- Include SHA256 checksums, SBOM, and GitHub Attestations for security
- Build workflow: `.github/workflows/build-binaries.yml`

---

## Office Application Support

### Supported Applications (Current)
- **Excel**: Full support (VBA modules, classes, forms)
- **Word**: Full support (VBA modules, classes, forms)
- **Access**: Basic support (modules and classes only, **NO forms** - Access forms handled differently)
- **PowerPoint**: Full support (VBA modules, classes, forms)

### Edit Mode Sync Behavior
- **Initial export**: All VBA code exported to files when `edit` command starts
- **One-way sync**: External editor → Office VBA (on save [CTRL+S])
- **NO bi-directional sync**: Changes made in Office VBA editor while edit mode is active are NOT synced back to external files
- **Future consideration**: Bi-directional sync may be added (e.g., via keyboard shortcut in terminal while edit mode active)

### VBA Trust Requirement
- All operations require "Trust access to VBA project object model" enabled
- `check` command verifies this setting
- Provide clear error messages when not enabled

---

## User Communication

### CLI Messages
- Use clear, actionable messages
- `IMPORTANT:` prefix for critical warnings
- `NOTE:` prefix for helpful information
- Include next steps when operations fail
- Be explicit about data loss risks (overwrites, deletions)

### User Prompts
- Use `--force-overwrite` flag to bypass safety prompts (for CI/CD)
- Always confirm before destructive operations (unless forced)
- Explain consequences clearly in prompts

---

## Security Practices

### Binary Verification
- Provide multiple verification methods: GitHub Attestations (preferred), SHA256 checksums, SBOM
- Document verification process in `SECURITY_VERIFICATION.md`
- Include checksums in two formats for compatibility

### Vulnerability Management
- Security policy documented in `SECURITY.md`
- Report vulnerabilities via GitHub Security Advisories
- Dependabot provides automated monitoring
- Response timeline: 48hr initial, 7-30 day resolution target
