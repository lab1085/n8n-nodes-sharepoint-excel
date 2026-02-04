# Release Process

This document describes the release workflow for n8n-nodes-sharepoint-excel.

## Branch Strategy

- **develop** - Active development branch
- **main** - Stable release branch, receives squash merges from develop

## Release Workflow

### 1. Prepare the Release

On `develop` branch:

```bash
# Update CHANGELOG.md with all changes since last release
# Update version in package.json
git add CHANGELOG.md package.json
git commit -m "chore: bump version to vX.Y.Z and update changelog"
git push origin develop
```

### 2. Create PR to Main

```bash
gh pr create --base main --head develop --title "Release vX.Y.Z" --body "$(cat <<'EOF'
## Summary
Release vX.Y.Z

## Changes
[Copy relevant section from CHANGELOG.md]
EOF
)"
```

Add label and assignee:

```bash
gh pr edit --add-assignee @me --add-label "release"
```

Then **squash merge** the PR via GitHub UI or CLI:

```bash
gh pr merge --squash
```

### 3. Create the Release

After the PR is merged, create the release from **main** branch:

```bash
gh release create vX.Y.Z --target main --title "vX.Y.Z" --notes "$(cat <<'EOF'
> YYYY-MM-DD

## Added

- Feature 1
- Feature 2

## Fixed

- Fix 1
- Fix 2
EOF
)"
```

This triggers the GitHub Action that publishes to npm automatically.

Release notes format:
- Date as blockquote at the top: `> 2026-02-02`
- Use `## Added`, `## Fixed`, `## Changed` headers
- Copy content from CHANGELOG.md

### 4. Sync Develop with Main

After the PR is merged:

```bash
git checkout develop
git pull origin main
git push origin develop
```

This merges the squash commit from main into develop, preventing conflicts on the next release.

## Version Numbering

Follow [Semantic Versioning](https://semver.org/):

- **MAJOR** (X.0.0) - Breaking changes
- **MINOR** (0.X.0) - New features, backwards compatible
- **PATCH** (0.0.X) - Bug fixes only

## Setup (One-Time)

### npm Token

1. Create an npm access token at https://www.npmjs.com/settings/lab1085/tokens
2. Add it as a GitHub secret named `NPM_TOKEN` in your repo settings

### Create Develop Branch

```bash
git checkout -b develop
git push -u origin develop
```

## Checklist

- [ ] CHANGELOG.md updated with all changes
- [ ] Version bumped in package.json
- [ ] Changes committed and pushed to develop
- [ ] PR created from develop to main
- [ ] PR labeled and assigned
- [ ] PR squash merged
- [ ] Release created from main (triggers npm publish)
- [ ] Develop synced with main
