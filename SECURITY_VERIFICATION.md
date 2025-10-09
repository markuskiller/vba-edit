# Security Verification Guide for vba-edit Binaries

## Overview

While vba-edit binaries are currently **unsigned**, we provide multiple security verification methods to ensure authenticity and integrity of downloads.

---

## 🔐 Verification Methods (Best to Good)

### 1. ✅ GitHub Attestations (RECOMMENDED - Cryptographic Proof)

**What it is**: GitHub cryptographically signs metadata proving the binary was built by our official GitHub Actions workflow.

**How to verify** (requires GitHub CLI):

```powershell
# Install GitHub CLI if needed
winget install GitHub.cli

# Verify the binary
gh attestation verify excel-vba.exe --owner markuskiller
```

**What you see if valid**:
```
✓ Verification succeeded!

Attestation verified for excel-vba.exe using build provenance
Repository: markuskiller/vba-edit
Workflow: Build and Attach Windows Binaries to Release
```

**Benefits**:
- ✅ Cryptographic proof from GitHub (not us)
- ✅ Verifies exact workflow that built it
- ✅ Confirms no tampering after build
- ✅ Works without code signing

---

### 2. ✅ SHA256 Checksum Verification

**What it is**: Verify the file hasn't been corrupted or tampered with during download.

**How to verify**:

```powershell
# Download SHA256SUMS from the same release
# Calculate hash of your download
Get-FileHash excel-vba.exe -Algorithm SHA256

# Compare the output with the corresponding line in SHA256SUMS
# They must match EXACTLY
```

**Benefits**:
- ✅ Detects download corruption
- ✅ Detects file tampering
- ✅ Standard industry practice

**Limitations**:
- ⚠️ Only works if you trust the download source (GitHub releases)
- ⚠️ Won't help if entire release was compromised

---

### 3. ✅ SBOM Review (Transparency)

**What it is**: Software Bill of Materials - a list of all dependencies included in the binary.

**How to use**:

1. Download `SBOM.txt` from release
2. Review the dependencies list
3. Check for any suspicious or unexpected packages

**Benefits**:
- ✅ Transparency about what's included
- ✅ Helps security audits
- ✅ Allows tracking of vulnerable dependencies

---

### 4. ✅ Build from Source (Ultimate Trust)

**What it is**: Build the binaries yourself from source code.

**How to build**:

```powershell
# Clone repository
git clone https://github.com/markuskiller/vba-edit.git
cd vba-edit

# Checkout specific release tag
git checkout v0.4.1

# Install dependencies
pip install uv
uv sync

# Build binaries
uv run python create_binaries.py

# Your binaries are in dist/
```

**Benefits**:
- ✅ Ultimate verification - you built it
- ✅ No need to trust our builds
- ✅ Can audit source code first

**Limitations**:
- ⚠️ Requires Python and build tools
- ⚠️ Takes more time

---

## 🚨 Security Warnings to Heed

### ⚠️ Download Source Matters

**ONLY download from**:
- ✅ Official GitHub releases: https://github.com/markuskiller/vba-edit/releases
- ✅ Official PyPI package: https://pypi.org/project/vba-edit/

**NEVER download from**:
- ❌ Third-party websites
- ❌ File sharing sites
- ❌ Email attachments
- ❌ Direct messages

### ⚠️ Windows SmartScreen Warnings

**Expected behavior**:
1. Windows shows: "Windows protected your PC"
2. Click "More info"
3. Click "Run anyway"

**Why this happens**:
- Binaries are unsigned (code signing costs money)
- Windows doesn't recognize new publishers
- This is NORMAL for unsigned open-source software

**When to be concerned**:
- ❌ If antivirus shows SPECIFIC threats (not just "unknown publisher")
- ❌ If checksums don't match
- ❌ If GitHub attestation fails
- ❌ If you didn't download from official GitHub releases

---

## 🔒 Security Best Practices

### For Regular Users:

1. **Use GitHub Attestations** (most reliable)
   ```powershell
   gh attestation verify excel-vba.exe --owner markuskiller
   ```

2. **Verify checksums** (backup verification)
   ```powershell
   Get-FileHash excel-vba.exe -Algorithm SHA256
   # Compare with SHA256SUMS
   ```

3. **Download only from GitHub releases**

### For Security-Conscious Users:

1. **Build from source** (ultimate trust)
2. **Review SBOM** for dependencies
3. **Check commit history** on GitHub
4. **Verify release signatures** (when we add GPG signing)

### For Enterprise/Regulated Environments:

1. **Build from source** using your own CI/CD
2. **Sign with your own certificate**
3. **Use pip installation** instead (uses Python's signed interpreter)
4. **Wait for code signing** (we're working on it)

---

## 🛡️ Our Security Commitment

### What We Do:

✅ **Open Source**: All code is public and auditable
✅ **Automated Builds**: Built by GitHub Actions (transparent, repeatable)
✅ **Attestations**: Cryptographic proof of build provenance
✅ **Checksums**: SHA256 for integrity verification
✅ **SBOM**: Full transparency of dependencies
✅ **Responsible Disclosure**: Security issues handled privately

### What We're Working On:

🔄 **Code Signing**: Applying for SignPath.io (free for open source)
🔄 **GPG Signatures**: For checksum files
🔄 **Security Policy**: SECURITY.md with vulnerability reporting process

---

## 📊 Comparison: Verification Strength

| Method | Strength | Ease | Requires |
|--------|----------|------|----------|
| GitHub Attestations | 🔒🔒🔒🔒🔒 | Easy | GitHub CLI |
| Build from Source | 🔒🔒🔒🔒🔒 | Hard | Dev tools |
| SHA256 + GitHub | 🔒🔒🔒 | Easy | None |
| Code Signing | 🔒🔒🔒🔒 | Easy | None (coming) |

---

## 🆘 Report Security Issues

**Found a security vulnerability?**

**DO NOT** open a public issue.

**DO** email: [Add your security contact email]

Or use GitHub Security Advisories:
https://github.com/markuskiller/vba-edit/security/advisories

---

## 📚 Additional Resources

- [GitHub Attestations Documentation](https://docs.github.com/en/actions/security-guides/using-artifact-attestations-to-establish-provenance-for-builds)
- [SBOM Explained](https://www.cisa.gov/sbom)
- [Supply Chain Security Best Practices](https://slsa.dev/)

---

**Last Updated**: October 9, 2025  
**Version**: 0.4.1  
**Status**: Unsigned binaries with attestations + checksums
