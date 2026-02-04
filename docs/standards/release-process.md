# Release Process

This document describes the release workflow for `@lab1095/n8n-nodes-sharepoint-excel`.

## Overview

Releases are published to npm using `npm run release`, which wraps `release-it`. The process handles version bumping, changelog generation, git tagging, GitHub releases, and npm publishing in one command.

## Prerequisites

- npm account with access to `@lab1095` org
- GitHub repo push access
- Logged into npm (`npm login`)

## Branch Strategy

- **main** - Stable release branch
- **feature branches** - For development work, squash merged to main

## Release Workflow

### 1. Prepare for Release

Ensure all changes are merged to `main` and tested.

### 2. Run the Release

```bash
npm login  # if session expired (tokens last ~12 hours)
npm run release
```

This will:

1. Prompt for version increment (patch/minor/major)
2. Run `npm run lint && npm run build`
3. Update `CHANGELOG.md` via auto-changelog
4. Bump version in `package.json`
5. Create git commit and tag
6. Push to GitHub
7. Create GitHub release (opens browser if no GITHUB_TOKEN)
8. Publish to npm

## Configuration

### `.npmrc`

```
access=public
```

Required for scoped packages (`@lab1095/*`) to publish as public.

### Commit Messages

The release commit must follow conventional commits format. The `n8n-node release` creates commits like `Release X.Y.Z`, but commitlint requires a type prefix. If the commit fails, the package may still be published to npm. Fix manually:

```bash
# Check npm for published version
npm view @lab1095/n8n-nodes-sharepoint-excel version

# Update package.json to match, then:
git add package.json CHANGELOG.md
git commit -m "chore(release): X.Y.Z"
git tag vX.Y.Z
git push origin main --tags

# Create GitHub release
gh release create vX.Y.Z --target main --title "vX.Y.Z" --notes "..."
```

## Version Numbering

Follow [Semantic Versioning](https://semver.org/):

- **MAJOR** (X.0.0) - Breaking changes
- **MINOR** (0.X.0) - New features, backwards compatible
- **PATCH** (0.0.X) - Bug fixes only

## Troubleshooting

### "Access token expired or revoked"

```bash
npm login
```

### "402 Payment Required - private packages"

Ensure `.npmrc` has `access=public`.

### Commitlint rejects "Release X.Y.Z"

See "Commit Messages" section above for manual fix.

## Checklist

- [ ] All changes merged to main
- [ ] Logged into npm
- [ ] `npm run release` completed successfully
- [ ] GitHub release created
