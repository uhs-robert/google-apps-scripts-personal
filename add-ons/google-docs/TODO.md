# TODO

## Manifest Configuration

The current `appsscript.json` includes an `addOns` block with `onFileScopeGrantedTrigger`:

```json
"addOns": {
  "docs": {
    "onFileScopeGrantedTrigger": {
      "runFunction": "onOpen"
    }
  }
}
```

**Issue:** According to Google's official documentation, `onFileScopeGrantedTrigger` is designed for CardService-based Google Workspace add-ons that need the `drive.file` scope for REST API access. It's not the correct approach for menu-based Editor add-ons.

**Current Status:** This configuration works in practice, so no immediate change needed.

**Future Consideration:** For alignment with best practices documented in README.md, consider:
- Removing the entire `addOns` block (the `onOpen()` simple trigger is automatically detected)
- OR keeping it if planning to migrate to a CardService-based interface

See README.md for the recommended simplified manifest structure for menu-based Editor add-ons.

## Code.js

- TODO: Change `createMenu()` to `createAddonMenu()` when publishing as an add-on (see comment in Code.js:3-4)
