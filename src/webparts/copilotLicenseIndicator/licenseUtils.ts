/**
 * Microsoft 365 Copilot の SKU ID
 * Graph API から取得されるライセンス一覧と照合するために使用する
 */
const COPILOT_SKU_ID = '639dec6b-bb19-468b-871c-c5c441c4b0cb';

/**
 * ライセンスエントリを表すインターフェース
 */
export interface ILicenseEntry {
  /** ライセンスのSKU ID */
  skuId: string;
  /** ライセンスのSKU パーツ番号 */
  skuPartNumber: string;
}

/**
 * ライセンス確認結果を表すインターフェース
 */
export interface ILicenseCheckResult {
  /** Copilot ライセンスが見つかったかどうか */
  found: boolean;
  /** 見つかったライセンスの名前。見つからない場合は空文字列 */
  name: string;
}

/**
 * ライセンス一覧の中に Microsoft 365 Copilot の SKU が含まれるか確認する
 * @param licenses - 確認対象のライセンスエントリ配列
 * @returns ライセンス確認結果。Copilot SKU が存在する場合は found=true と SKU パーツ番号を返す
 */
export function findCopilotLicense(licenses: ILicenseEntry[]): ILicenseCheckResult {
  // ライセンス配列から Copilot SKU ID に一致するエントリを検索する
  const match = licenses.find(l => l.skuId === COPILOT_SKU_ID);
  return match
    ? { found: true, name: match.skuPartNumber }
    : { found: false, name: '' };
}
