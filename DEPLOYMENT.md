# SendMeBot deployment targets

There is intentionally no tracked `.clasp.json`. A bare `clasp push` must fail.

## Local setup

1. Copy `config/targets.example.json` to ignored `config/targets.local.json`.
2. Keep each real clasp profile ignored and set its `rootDir`:
   - bound and template profiles: repository root;
   - add-on development profile at `installer/.clasp.development.json`: `dist/development`;
   - add-on production profile at `installer/.clasp.production.json`: `dist/production`.
3. Copy `config/environments.example.json` to ignored `config/environments.local.json` and enter the versioned template Spreadsheet IDs.

Inspect a target without changing it:

```sh
node scripts/clasp-target.mjs status bound-dev --dry-run
```

Push development only after tests pass:

```sh
npm test
node scripts/clasp-target.mjs push bound-dev
```

Add-on pushes automatically rebuild the environment-specific generated artifact before clasp runs.

## Production guard

Production pushes require all of the following:

- an explicit production alias;
- a clean Git worktree;
- an exact Git tag on the checked-out commit;
- `--confirm-production=<alias>`.

Example:

```sh
node scripts/clasp-target.mjs push bound-prod --confirm-production=bound-prod
```

Before the first Akamai production push, seed `SENDMEBOT_ENVIRONMENT_CONFIG` in Script Properties with the current tracker defaults, colors, logo, and assistant URL. Archive the frozen production source and complete the development smoke-test matrix first.
