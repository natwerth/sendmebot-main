# SendMeBot release metadata

The published metadata endpoint contains policy only. Executable Apps Script source is generated from the canonical repository and embedded in the corresponding installer build.

Required fields:

```json
{
  "channel": "stable",
  "version": "1.0.0",
  "payloadSha256": "base64url sha256",
  "minUpdaterVersion": "1.0.0",
  "updateMode": "automatic",
  "requiresReauthorization": false,
  "rolloutPercent": 100
}
```

Automatic updates run only when the endpoint version and hash exactly match the embedded payload. Releases adding OAuth scopes must set `requiresReauthorization` and use a manual update path.

Production release order:

1. Run all tests from a clean commit.
2. Build and verify the canonical payload hash.
3. Create an immutable sanitized template version.
4. Tag the exact commit.
5. Build the production installer with ignored local environment configuration.
6. Update metadata only after the matching add-on release is available.
