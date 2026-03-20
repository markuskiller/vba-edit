# GitHub Actions SHA Pinning Reference

**Last Updated**: 2026-03-20  
**Branch**: dev (synced with main)
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
- **Current Pin**: `de0fac2e4500dabe0009e67214ff5f5447ce83dd`
- **Version**: v6.0.2
- **Release URL**: https://github.com/actions/checkout/releases/tag/v6.0.2
- **Used in**: All workflows (build-binaries.yml, test.yaml, lint.yaml, publish.yaml)

### actions/setup-python
- **Current Pin**: `a309ff8b426b58ec0e2a45f0f869d46889d02405`
- **Version**: v6.2.0
- **Release URL**: https://github.com/actions/setup-python/releases/tag/v6.2.0
- **Used in**: All workflows (build-binaries.yml, test.yaml, lint.yaml, publish.yaml)

### astral-sh/setup-uv
- **Current Pin**: `e06108dd0aef18192324c70427afc47652e63a82`
- **Version**: v7.5.0
- **Release URL**: https://github.com/astral-sh/setup-uv/releases/tag/v7.5.0
- **Used in**: build-binaries.yml only
- **Notable change**: Uses `astral-sh/versions` as version provider (no more GitHub API rate-limits on version resolution)

### actions/attest-build-provenance
- **Current Pin**: `a2bbfa25375fe432b6a289bc6b6cd05ecd0c4c32`
- **Version**: v4.1.0
- **Release URL**: https://github.com/actions/attest-build-provenance/releases/tag/v4.1.0
- **Used in**: build-binaries.yml only

### actions/upload-artifact
- **Current Pin**: `bbbca2ddaa5d8feaa63e36b76fdaad77386f024f`
- **Version**: v7.0.0
- **Release URL**: https://github.com/actions/upload-artifact/releases/tag/v7.0.0
- **Used in**: build-binaries.yml only

### softprops/action-gh-release
- **Current Pin**: `153bb8e04406b158c6c84fc1615b65b24149a1fe`
- **Version**: v2.6.1
- **Release URL**: https://github.com/softprops/action-gh-release/releases/tag/v2.6.1
- **Used in**: build-binaries.yml only
- **Notable change**: Fixes discussion category preservation on publish; recovery of concurrent asset metadata 404s

## Updating Pinned Actions

When updating to a new version:

1. **Check the Release Page**: Visit the action's releases page
2. **Find the Commit SHA**: Look for the commit hash on the release (e.g., "Commit abc1234")
3. **Update the Workflow**: Replace both the SHA and the comment
4. **Test Thoroughly**: Run the workflow to ensure compatibility

### Example Update Process

```yaml
# Before
uses: actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683 # v5 (v4.2.2)

# After updating to v4.3.0
uses: actions/checkout@NEW_COMMIT_SHA_HERE # v5 (v4.3.0)
```

## Comment Format Convention

We use the following comment format for clarity:

```yaml
uses: owner/action@COMMIT_SHA # vX.Y.Z
```

Where `vX.Y.Z` is the actual release version the SHA points to (e.g., `v5.0.0`, `v6.0.0`).

This makes it clear exactly which version is pinned.

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
