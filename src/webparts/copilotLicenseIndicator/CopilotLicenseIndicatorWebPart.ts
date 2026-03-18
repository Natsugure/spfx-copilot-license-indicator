import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { findCopilotLicense, ILicenseEntry } from './licenseUtils';

/**
 * Web パーツのプロパティを定義するインターフェース。
 * 現時点ではカスタムプロパティを持たない。
 */
// eslint-disable-next-line @typescript-eslint/no-empty-interface
export interface ICopilotLicenseIndicatorWebPartProps {}

/**
 * Microsoft Graph `/me/licenseDetails` エンドポイントのレスポンス形式を表すインターフェース。
 */
interface ILicenseDetailsResponse {
  /** ライセンスエントリの配列 */
  value: ILicenseEntry[];
}

/**
 * Copilot ライセンス状況をユーザーに表示する SPFx Web パーツ。
 *
 * Microsoft Graph API を使用してサインイン済みユーザーのライセンス情報を取得し、
 * Microsoft 365 Copilot ライセンスの有無を Fluent UI コンポーネントで表示する。
 */
export default class CopilotLicenseIndicatorWebPart extends BaseClientSideWebPart<ICopilotLicenseIndicatorWebPartProps> {

  /**
   * ユーザーが Copilot ライセンスを保有しているかどうかを示すフラグ。
   * `onInit()` でライセンス取得後に更新される。
   */
  private _hasLicense: boolean = false;

  /**
   * 検出された Copilot ライセンスの表示名。
   * ライセンスが存在する場合に `onInit()` で設定される。
   */
  private _licenseName: string = '';

  /**
   * ライセンス情報の取得中であるかどうかを示すフラグ。
   * 初期値は `true` で、取得完了後に `false` へ変わる。
   */
  private _isLoading: boolean = true;

  /**
   * ライセンス取得時にエラーが発生した場合のメッセージ。
   * エラーが存在しない場合は空文字列。
   */
  private _errorMessage: string = '';

  /**
   * React のレンダリング先 DOM 要素。
   * `render()` メソッドの中で使用され、`onDispose()` でアンマウントに用いる。
   * 初回の `render()` 呼び出し時に `domElement` が設定され、以降は再利用される。
   */
  private _reactRoot: Element | undefined;

  /**
   * Web パーツの UI をレンダリングする。
   *
   * 初回呼び出し時に React ルートを遅延設定し、以降は同じ DOM 要素を再利用する。
   * `_isLoading` が `true` の場合は Spinner を表示し、
   * ライセンス取得完了後はその結果に応じた MessageBar を表示する。
   */
  public render(): void {
    if (!this._reactRoot) {
      this._reactRoot = this.domElement;
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

    ReactDOM.render(element, this._reactRoot);
  }

  /**
   * Web パーツの初期化処理。
   *
   * Microsoft Graph API の `/me/licenseDetails` エンドポイントを呼び出し、
   * Copilot ライセンスの有無を確認する。
   * 取得完了後（成功・失敗いずれの場合も）に `_isLoading` を `false` に設定して
   * `render()` を再度呼び出し、結果を UI に反映させる。
   *
   * @returns 初期化完了を表す Promise
   */
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

  /**
   * Web パーツが破棄される際のクリーンアップ処理。
   *
   * React ルートが存在する場合はアンマウントしてメモリリークを防ぐ。
   */
  protected onDispose(): void {
    if (this._reactRoot) {
      ReactDOM.unmountComponentAtNode(this._reactRoot);
    }
  }

  /**
   * この Web パーツのデータバージョンを返す。
   *
   * プロパティのシリアライズ互換性管理に使用される。
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * プロパティペインの設定を返す。
   *
   * この Web パーツはカスタムプロパティを持たないため、空のグループを返す。
   *
   * @returns プロパティペインの設定オブジェクト
   */
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
