# copilot-license-indicator

## Summary

SharePoint Framework (SPFx) の Web パーツとして動作し、サインインユーザーが **Microsoft 365 Copilot** ライセンスを保持しているかどうかを Fluent UI コンポーネントで表示します。

Microsoft Graph API の `/me/licenseDetails` エンドポイントを呼び出し、Copilot の SKU ID (`639dec6b-bb19-468b-871c-c5c441c4b0cb`) に一致するライセンスの有無を確認します。

- ライセンスあり → 緑の MessageBar にライセンス名を表示
- ライセンスなし → 黄色の MessageBar で通知
- エラー発生時 → 赤の MessageBar にエラーメッセージを表示
- 読み込み中 → Spinner を表示

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

- Node.js >= 22.14.0 < 23.0.0
- pnpm (推奨)
- Microsoft 365 テナント（開発環境）

## Solution

| Solution                  | Author(s) |
| ------------------------- | --------- |
| copilot-license-indicator | Natsugure   |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 0.0.1   | March 19, 2026   | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

```bash
# リポジトリをクローン
git clone <repository-url>
cd copilot-license-indicator

# 依存関係のインストール
pnpm install

# 開発サーバーの起動
heft start --clean
```

ブラウザで `https://localhost:4321/temp/workbench.html` を開き、Web パーツを追加して動作を確認できます。

## Build & Deploy

```bash
# テスト・ビルド・パッケージング（.sppkg 生成）
heft test --clean --production && heft package-solution --production
```

生成された `sharepoint/solution/*.sppkg` をアプリカタログにアップロードしてください。

## Features

- Microsoft Graph API を使用してサインインユーザーのライセンス情報を取得
- Microsoft 365 Copilot SKU ID による正確なライセンス判定
- Fluent UI (`@fluentui/react`) の MessageBar / Spinner でステート別 UI を表示
- SPFx 1.22.2 / React 17 / TypeScript 5.8 で実装
- SharePoint Web パーツ・Teams タブ・Teams Personal App のホストに対応

## Project Structure

```
src/
└── webparts/
    └── copilotLicenseIndicator/
        ├── CopilotLicenseIndicatorWebPart.ts   # Web パーツ本体
        ├── licenseUtils.ts                      # ライセンス判定ロジック
        └── tests/
            └── licenseUtils.test.ts             # ユニットテスト
```

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Fluent UI React](https://developer.microsoft.com/en-us/fluentui#/controls/web)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp)
- [Heft Documentation](https://heft.rushstack.io/)
