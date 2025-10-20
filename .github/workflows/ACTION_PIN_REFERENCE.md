# GitHub Actions SHA Pinning Reference

**Last Updated**: 2025-10-20  
**Branch**: main  
**Reason**: Security best practice - pinning actions to commit SHAs prevents supply chain attacks

## Why Pin to Commit SHAs?

Pinning third-party GitHub Actions to full commit SHAs is currently the **only way** to use an action as an immutable release. This security practice helps mitigate the risk of a bad actor adding a backdoor to an action's repository.

**Security Benefits**:
- Prevents automatic updates to potentially compromised versions
- Requires explicit review before updating to new versions
- Makes it harder for attackers to exploit tag/branch manipulation
- Provides cryptographic verification through Git's SHA-1 hashing

## Pinned Actions Reference

### actions/checkout
- **Current Pin**: `08c6903cd8c0fde910a37f88322edcfb5dd907a8`
- **Version**: v5.0.0
- **Release Date**: August 11, 2024
- **Release URL**: https://github.com/actions/checkout/releases/tag/v5.0.0
- **Used in**: All workflows (build-binaries.yml, test.yaml, lint.yaml, publish.yaml)
- **Breaking Change**: Requires runner v2.327.1+ (uses node24)

### actions/setup-python
- **Current Pin**: `e797f83bcb11b83ae66e0230d6156d7c80228e7c`
- **Version**: v6.0.0
- **Release Date**: September 4, 2024
- **Release URL**: https://github.com/actions/setup-python/releases/tag/v6.0.0
- **Used in**: All workflows (build-binaries.yml, test.yaml, lint.yaml, publish.yaml)
- **Breaking Change**: Requires runner v2.327.1+ (uses node24)

### astral-sh/setup-uv
- **Current Pin**: `eb1897b8dc4b5d5bfe39a428a8f2304605e0983c`
- **Version**: v7.0.0
- **Release Date**: October 2025 (2 weeks ago as of 2025-10-20)
- **Release URL**: https://github.com/astral-sh/setup-uv/releases/tag/v7.0.0
- **Used in**: build-binaries.yml only
- **Breaking Changes**: Uses node24 instead of node20, removed deprecated `server-url` input

### actions/attest-build-provenance
- **Current Pin**: `977bb373ede98d70efdf65b84cb5f73e068dcc2a`
- **Version**: v3.0.0
- **Release Date**: August 28, 2024
- **Release URL**: https://github.com/actions/attest-build-provenance/releases/tag/v3.0.0
- **Used in**: build-binaries.yml only

### actions/upload-artifact
- **Current Pin**: `c7d193f32edcb7bfad88892161225aeda64e9392`
- **Version**: v4.0.0
- **Release Date**: December 14, 2023
- **Release URL**: https://github.com/actions/upload-artifact/releases/tag/v4.0.0
- **Used in**: build-binaries.yml only

### softprops/action-gh-release
- **Current Pin**: `a6c7483a42ee9d5daced968f6c217562cd680f7f`
- **Version**: v2.0.0
- **Release Date**: March 8, 2024
- **Release URL**: https://github.com/softprops/action-gh-release/releases/tag/v2.0.0
- **Used in**: build-binaries.yml only

## Updating Pinned Actions

When updating to a new version:

1. **Check the Release Page**: Visit the action's releases page
2. **Find the Commit SHA**: Look for the commit hash on the release (e.g., "Commit abc1234")
3. **Update the Workflow**: Replace both the SHA and the comment
4. **Test Thoroughly**: Run the workflow to ensure compatibility

### Example Update Process

```yaml
# Before
uses: actions/checkout@08c6903cd8c0fde910a37f88322edcfb5dd907a8 # v5.0.0

# After updating to v5.1.0
uses: actions/checkout@NEW_COMMIT_SHA_HERE # v5.1.0
```

## Comment Format Convention

We use the following comment format for clarity:

```yaml
uses: owner/action@COMMIT_SHA # vX.Y.Z
```

Where `vX.Y.Z` is the actual release version the SHA points to (e.g., `v5.0.0`, `v6.0.0`).

This makes it clear exactly which version is pinned.

## Branch-Specific Differences

### Main Branch vs Dev Branch
- **setup-uv**: Main uses v7.0.0, dev may use v4.0.0
  - Main: `eb1897b8dc4b5d5bfe39a428a8f2304605e0983c` (v7.0.0)
  - Dev: `d8db0a86d3d88f3017a4e6b8a1e2b234e7a0a1b5` (v4.0.0)

### paths-ignore Configuration
- **test.yaml**: Main branch includes additional ignore patterns
  - Ignores: `CHANGELOG.md`, `README.md`, `AUTHORS.md`, `.github/**` (except workflows)
- **lint.yaml**: Main branch includes additional ignore patterns
  - Ignores: `CHANGELOG.md`, `README.md`, `AUTHORS.md`, `.github/**` (except workflows)

## Automated Dependency Updates

**Note**: Standard Dependabot does NOT support SHA pinning for GitHub Actions. Consider:
- **Manual monthly reviews** of action releases
- **Security-focused updates**: Priority on security patches
- **GitHub Advanced Security**: May provide SHA update suggestions

## References

- [GitHub Actions Security Hardening](https://docs.github.com/en/actions/security-guides/security-hardening-for-github-actions#using-third-party-actions)
- [Pinning Actions to Commit SHAs](https://docs.github.com/en/actions/security-guides/security-hardening-for-github-actions#using-third-party-actions)
- OpenGrep Rule: `yaml.github-actions.security.third-party-action-not-pinned-to-commit-sha`

---

**Maintainer Notes**:
- Check for action updates quarterly or when security advisories are published
- Always review release notes before updating
- Test updated actions in a feature branch first
- Document breaking changes in CHANGELOG.md
- Keep main and dev branches in sync for security updates
