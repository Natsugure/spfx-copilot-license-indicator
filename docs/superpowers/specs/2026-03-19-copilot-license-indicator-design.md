# Copilot License Indicator Web Part — Design Spec

**Date:** 2026-03-19
**SPFx Version:** 1.22.2
**Status:** Approved

---

## Overview

A SharePoint Framework (SPFx) Web Part that checks whether the currently signed-in user has a Microsoft 365 Copilot license and displays the result using Fluent UI React components.

**Target SKU ID:** `639dec6b-bb19-468b-871c-c5c441c4b0cb`

---

## Architecture

### Data Flow

```
SPFx calls render() automatically (before onInit resolves)
  └─> render() shows Spinner (_isLoading=true)

onInit() runs asynchronously:
  └─> MSGraphClientFactory.getClient('3')
        └─> GET /me/licenseDetails
              └─> Check if any entry matches skuId `639dec6b-bb19-468b-871c-c5c441c4b0cb`
                    ├─ Match found → _hasLicense=true, _licenseName=<skuPartNumber>
                    └─ No match   → _hasLicense=false
              └─> set _isLoading=false
              └─> this.render() called explicitly
                    ├─ Error  → MessageBar (error, red)
                    ├─ Found  → MessageBar (success, green) + license name
                    └─ Not found → MessageBar (warning, yellow)
```

**`render()` / `onInit()` lifecycle contract:**
- SPFx automatically calls `render()` once before `onInit()` resolves. At that point `_isLoading === true`, so the Spinner is displayed.
- `onInit()` must explicitly call `this.render()` at the end after updating state, because SPFx does NOT re-call `render()` automatically after `onInit()` completes.
- The `render()` method is a pure function of instance state and safe to call at any time.

**Note on pagination:** `/me/licenseDetails` may return paginated results via `@odata.nextLink`. For this indicator, pagination is out of scope — only the first page of results is checked. This is sufficient for the vast majority of users.

### File Changes

| File | Change |
|------|--------|
| `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.ts` | Graph call, state management, React 17 `ReactDOM.render` rendering |
| `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.module.scss` | Minimal layout styles |
| `config/package-solution.json` | Add `webApiPermissionRequests` inside the `"solution"` object |

---

## Component Design

### Key Imports

```typescript
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
```

SPFx 1.22.2 ships with **React 17.0.1** (not React 18) as a shared library. `react` and `react-dom` should be listed in `devDependencies` only (SPFx provides them as externals at runtime). `@fluentui/react` must be in `dependencies`.

### State Fields

```typescript
private _hasLicense: boolean = false;
private _licenseName: string = '';
private _isLoading: boolean = true;
private _errorMessage: string = '';
private _reactRoot: Element | undefined;
```

`_reactRoot` stores the container DOM element (`this.domElement`) passed to `ReactDOM.render()` and `ReactDOM.unmountComponentAtNode()`.

### Rendering Logic

`render()` uses React 17's `ReactDOM.render()`. On the first call it sets `_reactRoot = this.domElement`. On subsequent calls it reuses the stored container. It calls `ReactDOM.render(element, this._reactRoot)` with the appropriate Fluent UI component.

| State | Fluent UI Component | Appearance |
|-------|-------------------|------------|
| `_isLoading === true` | `Spinner` (SpinnerSize.medium) | Spinner with "確認中..." label |
| `_errorMessage !== ''` | `MessageBar` (MessageBarType.error) | Red bar with error message |
| `_hasLicense === true` | `MessageBar` (MessageBarType.success) | Green bar + license name |
| `_hasLicense === false` | `MessageBar` (MessageBarType.warning) | Yellow bar |

### Graph API

- **Endpoint:** `GET /me/licenseDetails`
- **Required Permission:** `User.Read` (delegated) — sufficient for reading the signed-in user's own license details
- **Client:** `this.context.msGraphClientFactory.getClient('3')` — SPFx built-in, no extra packages needed
- **Type:** `MSGraphClientV3` from `@microsoft/sp-http`

### Permission Configuration (`package-solution.json`)

Add inside the `"solution"` object:

```json
"webApiPermissionRequests": [
  {
    "resource": "Microsoft Graph",
    "scope": "User.Read"
  }
]
```

**Important:** This array must be placed inside `"solution": { ... }`, not at the top level of the JSON file. After deploying the `.sppkg`, a tenant administrator must approve this permission request in the SharePoint Admin Center.

### Cleanup (`onDispose`)

Override `onDispose()` to unmount the React root and avoid memory leaks:

```typescript
protected onDispose(): void {
  if (this._reactRoot) {
    ReactDOM.unmountComponentAtNode(this._reactRoot);
  }
}
```

### Property Pane

The default `description` field is removed. The property pane is left as a minimal empty configuration (no user-configurable properties needed).

---

## Error Handling

- Graph API call failure (permission not yet approved, network error): catch the error, set `_errorMessage`, call `this.render()` to display `MessageBar` with `MessageBarType.error`.
- Loading state shown via `Spinner` while the async Graph call is in flight.

---

## Out of Scope

- Pagination handling for `/me/licenseDetails` (first page only)
- Displaying licenses beyond Copilot
- Admin-level license views (delegated user view only)
