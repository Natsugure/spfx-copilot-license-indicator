# Copilot License Indicator Web Part — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Modify the existing SPFx Web Part to display whether the signed-in user has a Microsoft 365 Copilot license by calling Microsoft Graph `/me/licenseDetails`.

**Architecture:** The Web Part calls the Graph API via SPFx's built-in `MSGraphClientFactory` on initialization, checks for SKU ID `639dec6b-bb19-468b-871c-c5c441c4b0cb`, and renders a Fluent UI `MessageBar` (success/warning/error) or `Spinner` using React 18's `createRoot` API. License-checking logic is extracted into a pure function (`licenseUtils.ts`) for testability.

**Tech Stack:** SPFx 1.22.2, TypeScript, React 18 (provided by SPFx), Fluent UI v8 (`@fluentui/react`), Microsoft Graph API, Jest via `@rushstack/heft`

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `src/webparts/copilotLicenseIndicator/licenseUtils.ts` | **Create** | Pure function: checks if a Copilot SKU exists in a license list |
| `src/webparts/copilotLicenseIndicator/tests/licenseUtils.test.ts` | **Create** | Unit tests for the license-checking logic |
| `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.ts` | **Modify** | Web Part class: Graph call, state, React rendering |
| `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.module.scss` | **Modify** | Remove unused scaffold styles |
| `config/package-solution.json` | **Modify** | Add `webApiPermissionRequests` for `User.Read` |

---

## Task 1: Add Graph API permission

**Files:**
- Modify: `config/package-solution.json`

- [ ] **Step 1: Add `webApiPermissionRequests` inside the `"solution"` object**

Open `config/package-solution.json`. Inside `"solution": { ... }`, add the following key (at the same level as `"name"`, `"id"`, etc.):

```json
"webApiPermissionRequests": [
  {
    "resource": "Microsoft Graph",
    "scope": "User.Read"
  }
]
```

The final `"solution"` object should look like:

```json
"solution": {
  "name": "copilot-license-indicator-client-side-solution",
  "id": "dc4f992e-ada9-4917-ada3-e96d6ab3200d",
  "version": "1.0.0.0",
  "includeClientSideAssets": true,
  "skipFeatureDeployment": true,
  "isDomainIsolated": false,
  "webApiPermissionRequests": [
    {
      "resource": "Microsoft Graph",
      "scope": "User.Read"
    }
  ],
  ...
}
```

- [ ] **Step 2: Commit**

```bash
git add config/package-solution.json
git commit -m "config: add User.Read Graph permission request"
```

---

## Task 2: Implement license-checking utility (TDD)

**Files:**
- Create: `src/webparts/copilotLicenseIndicator/licenseUtils.ts`
- Create: `src/webparts/copilotLicenseIndicator/tests/licenseUtils.test.ts`

- [ ] **Step 1: Write the failing tests**

Create `src/webparts/copilotLicenseIndicator/tests/licenseUtils.test.ts`:

```typescript
import { findCopilotLicense, ILicenseEntry } from '../licenseUtils';

const COPILOT_SKU_ID = '639dec6b-bb19-468b-871c-c5c441c4b0cb';

describe('findCopilotLicense', () => {
  it('returns found=true and the license name when Copilot SKU is present', () => {
    const licenses: ILicenseEntry[] = [
      { skuId: COPILOT_SKU_ID, skuPartNumber: 'Microsoft_365_Copilot' }
    ];
    const result = findCopilotLicense(licenses);
    expect(result.found).toBe(true);
    expect(result.name).toBe('Microsoft_365_Copilot');
  });

  it('returns found=false when Copilot SKU is not present', () => {
    const licenses: ILicenseEntry[] = [
      { skuId: 'other-sku-id', skuPartNumber: 'SOME_OTHER_LICENSE' }
    ];
    const result = findCopilotLicense(licenses);
    expect(result.found).toBe(false);
    expect(result.name).toBe('');
  });

  it('returns found=false for an empty license list', () => {
    const result = findCopilotLicense([]);
    expect(result.found).toBe(false);
    expect(result.name).toBe('');
  });

  it('finds Copilot SKU among multiple licenses', () => {
    const licenses: ILicenseEntry[] = [
      { skuId: 'other-sku-1', skuPartNumber: 'OTHER_1' },
      { skuId: COPILOT_SKU_ID, skuPartNumber: 'Microsoft_365_Copilot' },
      { skuId: 'other-sku-2', skuPartNumber: 'OTHER_2' }
    ];
    const result = findCopilotLicense(licenses);
    expect(result.found).toBe(true);
    expect(result.name).toBe('Microsoft_365_Copilot');
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
cd /Users/akizora1023/VSCode/study/microsoft/copilot-license-indicator
pnpm heft test --clean
```

Expected: build fails or tests fail with "Cannot find module '../licenseUtils'"

- [ ] **Step 3: Implement `licenseUtils.ts`**

Create `src/webparts/copilotLicenseIndicator/licenseUtils.ts`:

```typescript
const COPILOT_SKU_ID = '639dec6b-bb19-468b-871c-c5c441c4b0cb';

export interface ILicenseEntry {
  skuId: string;
  skuPartNumber: string;
}

export interface ILicenseCheckResult {
  found: boolean;
  name: string;
}

export function findCopilotLicense(licenses: ILicenseEntry[]): ILicenseCheckResult {
  const match = licenses.find(l => l.skuId === COPILOT_SKU_ID);
  return match
    ? { found: true, name: match.skuPartNumber }
    : { found: false, name: '' };
}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
pnpm heft test --clean
```

Expected: all 4 tests in `licenseUtils.test.ts` PASS

- [ ] **Step 5: Commit**

```bash
git add src/webparts/copilotLicenseIndicator/licenseUtils.ts
git add src/webparts/copilotLicenseIndicator/tests/licenseUtils.test.ts
git commit -m "feat: add licenseUtils with Copilot SKU check"
```

---

## Task 3: Implement Web Part rendering

**Files:**
- Modify: `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.ts`

- [ ] **Step 1: Replace the entire Web Part file**

Replace `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.ts` with:

```typescript
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as React from 'react';
import { createRoot, Root } from 'react-dom/client';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { findCopilotLicense, ILicenseEntry } from './licenseUtils';

// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ICopilotLicenseIndicatorWebPartProps {}

interface ILicenseDetailsResponse {
  value: ILicenseEntry[];
}

export default class CopilotLicenseIndicatorWebPart extends BaseClientSideWebPart<ICopilotLicenseIndicatorWebPartProps> {

  private _hasLicense: boolean = false;
  private _licenseName: string = '';
  private _isLoading: boolean = true;
  private _errorMessage: string = '';
  private _reactRoot: Root | undefined;

  public render(): void {
    if (!this._reactRoot) {
      this._reactRoot = createRoot(this.domElement);
    }

    let element: React.ReactElement;

    if (this._isLoading) {
      element = React.createElement(Spinner, {
        size: SpinnerSize.medium,
        label: '確認中...'
      });
    } else if (this._errorMessage) {
      element = React.createElement(
        MessageBar,
        { messageBarType: MessageBarType.error },
        this._errorMessage
      );
    } else if (this._hasLicense) {
      element = React.createElement(
        MessageBar,
        { messageBarType: MessageBarType.success },
        `Microsoft 365 Copilot ライセンスが付与されています (${this._licenseName})`
      );
    } else {
      element = React.createElement(
        MessageBar,
        { messageBarType: MessageBarType.warning },
        'Microsoft 365 Copilot ライセンスが付与されていません'
      );
    }

    this._reactRoot.render(element);
  }

  protected onInit(): Promise<void> {
    return this.context.msGraphClientFactory
      .getClient('3')
      .then((client: MSGraphClientV3) => {
        return client.api('/me/licenseDetails').get() as Promise<ILicenseDetailsResponse>;
      })
      .then((response: ILicenseDetailsResponse) => {
        const result = findCopilotLicense(response.value);
        this._hasLicense = result.found;
        this._licenseName = result.name;
      })
      .catch((error: Error) => {
        this._errorMessage = error.message || 'ライセンス情報の取得に失敗しました';
      })
      .then(() => {
        this._isLoading = false;
        this.render();
      });
  }

  protected onDispose(): void {
    if (this._reactRoot) {
      this._reactRoot.unmount();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: '' },
          groups: []
        }
      ]
    };
  }
}
```

- [ ] **Step 2: Build to verify no TypeScript errors**

```bash
pnpm heft test --clean
```

Expected: build succeeds, all tests pass

- [ ] **Step 3: Commit**

```bash
git add src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.ts
git commit -m "feat: implement Copilot license indicator Web Part"
```

---

## Task 4: Clean up SCSS

**Files:**
- Modify: `src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.module.scss`

- [ ] **Step 1: Replace with minimal styles**

Replace the entire file with:

```scss
@import '~@microsoft/sp-office-ui-fabric-core/dist/sass/SPFabricCore.scss';

.copilotLicenseIndicator {
  padding: 1em;
}
```

The Fluent UI `MessageBar` and `Spinner` components handle their own styling; no additional styles are needed.

- [ ] **Step 2: Build to verify no errors**

```bash
pnpm heft test --clean
```

Expected: build succeeds, all tests pass

- [ ] **Step 3: Commit**

```bash
git add src/webparts/copilotLicenseIndicator/CopilotLicenseIndicatorWebPart.module.scss
git commit -m "style: remove unused scaffold styles"
```

---

## Manual Verification (SharePoint)

After deploying to SharePoint:

1. Run `pnpm run build` to produce the `.sppkg` file at `sharepoint/solution/copilot-license-indicator.sppkg`
2. Upload the `.sppkg` to the SharePoint App Catalog
3. A tenant administrator must approve the `User.Read` permission in **SharePoint Admin Center → Advanced → API Access**
4. Add the Web Part to a SharePoint page
5. Verify:
   - If the current user **has** the Copilot license → green `MessageBar` with license name
   - If the current user **does not have** the license → yellow `MessageBar`
   - While loading → `Spinner` is displayed
