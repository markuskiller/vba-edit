# Security Policy

## Supported Versions

We take security seriously. The following versions of vba-edit currently receive security updates:

| Version | Supported          | Status | End of Support |
| ------- | ------------------ | ------ | -------------- |
| 0.4.x   | :white_check_mark: | Active development | Current |
| 0.3.x   | :warning:          | Critical fixes only | When 0.5.0 releases |
| < 0.3   | :x:                | No longer supported | - |

**Note**: As an alpha/beta project (0.x versions), we focus security support on the latest minor version. When 1.0 is released, we'll establish a longer-term support policy.

---

## Reporting a Vulnerability

**We take security vulnerabilities seriously.** If you discover a security issue, please help us address it responsibly.

### 🔒 Private Disclosure (Preferred)

**DO NOT** open a public GitHub issue for security vulnerabilities.

Please report security issues through one of these channels:

#### 1. GitHub Security Advisories (Recommended)
- Go to: https://github.com/markuskiller/vba-edit/security/advisories/new
- Click "Report a vulnerability"
- Fill in the details privately
- We'll respond within **48 hours**

#### 2. Email (Alternative)
- Email: m.killer@langui.ch
- Subject: `[SECURITY] vba-edit vulnerability`
- Include:
  - Description of the vulnerability
  - Steps to reproduce
  - Potential impact
  - Suggested fix (if any)

### ⏱️ What to Expect

**Response Timeline:**
- **Initial response**: Within 48 hours (usually faster)
- **Status update**: Every 3-5 days until resolved
- **Resolution target**: 
  - Critical: 7 days
  - High: 14 days
  - Medium: 30 days
  - Low: Next minor release

**Our Process:**
1. **Acknowledge** receipt of your report (48 hours)
2. **Investigate** and confirm the vulnerability (3-7 days)
3. **Develop** a fix in a private branch
4. **Test** the fix thoroughly
5. **Coordinate** disclosure timing with you
6. **Release** patched version
7. **Publish** security advisory with credit (if desired)

### 🎯 Scope

#### In Scope
Security vulnerabilities in:
- ✅ VBA code injection
- ✅ File path traversal
- ✅ Command injection
- ✅ Arbitrary code execution
- ✅ Information disclosure
- ✅ Privilege escalation
- ✅ Dependencies with known CVEs

#### Out of Scope
The following are **not** considered security vulnerabilities:
- ❌ Bugs without security impact
- ❌ Feature requests
- ❌ Issues requiring physical access
- ❌ Social engineering attacks
- ❌ DoS via malformed Office files (expected behavior - validate inputs)
- ❌ Windows SmartScreen warnings (binaries are unsigned by design)
- ❌ Antivirus false positives from heuristic scanners (PyInstaller executables are commonly flagged due to packing behavior, not malware)

### 🏆 Recognition

We believe in giving credit where it's due:

**If you report a valid security vulnerability:**
- ✅ Listed in security advisory (unless you prefer anonymity)
- ✅ Added to SECURITY.md Hall of Fame (optional)
- ✅ Mentioned in release notes
- ✅ Our sincere gratitude! 🙏

**We do NOT offer:**
- ❌ Bug bounties (we're an open-source project with no funding)
- ❌ Monetary rewards
- ❌ Swag or merchandise

---

## False Positives from Antivirus Software

### Why Antivirus Software Flags Our Executables

Our Windows binaries (.exe files) are built with PyInstaller, a legitimate tool that packages Python applications into standalone executables. Unfortunately, antivirus software frequently flags PyInstaller executables as potentially malicious due to:

1. **Packing/Bundling Behavior**: PyInstaller bundles the Python interpreter and libraries into a single executable, which looks similar to how malware packers work
2. **Self-Extracting Code**: The executable extracts Python runtime components at startup, triggering heuristic detection
3. **Code Obfuscation Appearance**: The packed structure can appear obfuscated to signature-based scanners
4. **Low Prevalence**: Unsigned executables from smaller projects lack the reputation scores that antivirus vendors use
5. **Generic Signatures**: Antivirus software uses broad heuristics that catch legitimate PyInstaller executables

### Common False Positive Detections

When scanning our executables with VirusTotal or similar services, you may see detections such as:
- `Trojan.Generic.*`
- `BehavesLike.Win64.Generic.*`
- `Trojan.Blank.Script.*`
- Generic "malicious" or "suspicious" flags

**Typical detection rate**: 4-10 detections out of 70+ scanners on VirusTotal.

**These are false positives.** The executables contain no malicious code.

### Why We Don't Code Sign (Yet)

Code signing certificates cost $300-500/year and require:
- Business verification process (weeks to months)
- Annual renewal fees
- Administrative overhead for a small open-source project

We're evaluating SignPath.io's free code signing for open-source projects (target: v0.6.0+), but this requires approval and setup time.

### How to Verify Our Executables Are Safe

Rather than relying solely on antivirus heuristics, use our security verification process:

1. **Build Attestations** (Strongest verification):
   - Every binary has a cryptographic attestation from GitHub
   - Proves the binary was built from our source code by GitHub Actions
   - Cannot be faked or tampered with
   - See SECURITY_VERIFICATION.md for instructions

2. **Verify Checksums**:
   - Download SHA256SUMS from the same release
   - Compare checksums to ensure file integrity
   - Detects any modification after build

3. **Review Source Code**:
   - All source code is public on GitHub
   - Review what the code actually does
   - Build the executable yourself if desired

4. **Check VirusTotal Results Carefully**:
   - **4-10 detections out of 70+ scanners is typical for PyInstaller executables**
   - Look at *which* vendors flag it (often lesser-known scanners)
   - Major vendors (Microsoft Defender, Norton, Kaspersky, etc.) typically don't flag it
   - Generic detection names indicate heuristic false positives, not actual malware

### Building From Source (Ultimate Verification)

If you're concerned about the pre-built binaries:

```bash
# Clone the repository
git clone https://github.com/markuskiller/vba-edit.git
cd vba-edit

# Install dependencies
pip install -e .[dev]

# Build binaries yourself
python create_binaries.py

# Your binaries are in ./dist/
```

Building from source gives you complete control and transparency.

### Reporting to Antivirus Vendors

If you encounter false positives, you can help by reporting them to antivirus vendors:
- **VirusTotal**: Use the "Report False Positive" option
- **Vendor-specific**: Submit to Microsoft, Norton, McAfee, etc.

Each report helps improve detection accuracy for all PyInstaller-based open-source tools.

### Our Commitment

- ✅ All builds happen on GitHub Actions (transparent, auditable)
- ✅ Source code is fully public and reviewable
- ✅ Build provenance is cryptographically verified (attestations)
- ✅ No obfuscation or packing beyond standard PyInstaller
- ✅ No telemetry, no network connections (except Office COM APIs)
- ✅ BSD-3-Clause licensed (permissive open source)

**Bottom line**: The detections are false positives caused by PyInstaller's legitimate behavior, not actual malware.

---

## Security Best Practices for Users

### When Using vba-edit

1. **Trust Your Office Files**
   - Only use vba-edit with Office files you trust
   - vba-edit accesses the VBA project model via COM automation
   - The tool imports/exports VBA code but does NOT execute it
   - Execution happens in Office VBA, outside vba-edit's control
   - Malicious VBA code (if imported) can harm your system when YOU run it in Office

2. **Validate Configuration Files**
   - Review `--conf` files before using them
   - Don't use configuration files from untrusted sources
   - Path traversal protections are in place, but validate anyway

3. **Keep vba-edit Updated**
   - Update regularly: `pip install -U vba-edit`
   - Check for security advisories in releases
   - Subscribe to GitHub releases for notifications

4. **Binary Verification**
   - See [SECURITY_VERIFICATION.md](./SECURITY_VERIFICATION.md) for how to verify binaries
   - Use GitHub Attestations: `gh attestation verify excel-vba.exe --owner markuskiller`
   - Verify checksums from releases

5. **Environment Security**
   - Run vba-edit in the context of the appropriate user
   - Don't run as administrator unless necessary
   - Be cautious when using `--force-overwrite` (skips safety prompts)

### When Developing with vba-edit

1. **Code Review**
   - Review VBA code before importing into Office
   - The tool trusts the `.bas/.cls` files you provide
   - Malicious code in those files will be imported

2. **Path Safety**
   - Use absolute paths or carefully validate relative paths
   - Don't construct paths from untrusted input
   - vba-edit validates paths, but be defensive

---

## Security Features

vba-edit includes several security features:

### ✅ Input Validation
- Path traversal protection
- File type validation
- Office file format verification

### ✅ Safe Defaults
- Confirmation prompts before overwriting files
- `--force-overwrite` required to skip safety checks
- No automatic code execution without explicit commands

### ✅ Transparency
- Open source - all code is auditable
- GitHub Attestations for binary provenance
- SBOM (Software Bill of Materials) included
- SHA256 checksums for integrity verification

### ✅ Dependency Management
- Regular dependency updates
- Security scanning via GitHub Dependabot
- Minimal dependencies to reduce attack surface

---

## Known Security Considerations

### Office VBA Model Access
**By design, vba-edit requires:**
- "Trust access to the VBA project object model" enabled in Office
- This is required for ANY programmatic VBA manipulation
- This setting is commonly used by VBA development tools

**Implications:**
- ⚠️ Other applications/macros can also access VBA projects when enabled
- ✅ You should disable this setting when not actively using VBA tools
- ✅ Run `excel-vba check` to verify the setting status

### Unsigned Windows Binaries
**Current state:**
- Windows executables are currently **unsigned**
- Users will see Windows SmartScreen warnings
- This is expected for open-source projects without funding

**Mitigation:**
- ✅ Use GitHub Attestations for verification
- ✅ Verify SHA256 checksums
- ✅ Build from source for ultimate trust
- 🔄 Code signing via SignPath.io planned

### COM Automation Limitations
**Known issues with Office COM:**
- Office applications may occasionally crash (Windows COM instability)
- This is a limitation of Microsoft's COM model, not vba-edit
- Always save your work before using `edit` mode

---

## Security Hall of Fame

We'd like to thank the following people for responsibly disclosing security issues:

<!-- When someone reports a vulnerability, add them here -->

*No security reports yet - be the first! 🎖️*

---

## Past Security Advisories

<!-- Security advisories will be listed here after publication -->

*No published security advisories yet.*

For a complete list, see: https://github.com/markuskiller/vba-edit/security/advisories

---

## Dependencies

We monitor our dependencies for known vulnerabilities:

- **GitHub Dependabot**: Automated security updates for dependencies
- **Dependency scanning**: Regular audits of dependency tree
- **Minimal dependencies**: Reducing attack surface

Current dependencies are documented in:
- `pyproject.toml` (source)
- `SBOM.txt` (included with binary releases)

---

## Questions?

If you have questions about our security policy that don't involve reporting a vulnerability:

- 💬 Open a public discussion: https://github.com/markuskiller/vba-edit/discussions
- 📧 Email: markus.killer@gmail.com
- 📝 Open a public issue: https://github.com/markuskiller/vba-edit/issues

---

**Last Updated**: October 9, 2025  
**Version**: 1.0  
**Project Status**: Alpha (0.4.x)
