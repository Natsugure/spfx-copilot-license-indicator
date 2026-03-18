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
