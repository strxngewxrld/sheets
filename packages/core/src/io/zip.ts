import { unzipSync, type Unzipped } from 'fflate';

/** Largest single part: V8 cannot hold a longer string anyway. */
export const MAX_ZIP_ENTRY_BYTES = 512 * 1024 * 1024;
export const MAX_ZIP_TOTAL_BYTES = 1024 * 1024 * 1024;
export const MAX_ZIP_ENTRIES = 20_000;

/**
 * Validate every central-directory entry before allocating decompressed data.
 * Stored entries are copied using their compressed size by fflate, so their
 * sizes must agree; otherwise a forged originalSize bypasses the budget.
 */
export function unzipLimited(bytes: Uint8Array): Unzipped {
  let total = 0;
  let count = 0;
  unzipSync(bytes, {
    filter: (file) => {
      if (file.compression === 0 && file.size !== file.originalSize) throw new Error('Not a valid xlsx file: inconsistent stored part size');
      total += file.originalSize;
      if (++count > MAX_ZIP_ENTRIES) throw new Error(`Not a valid xlsx file: more than ${MAX_ZIP_ENTRIES} parts`);
      if (file.originalSize > MAX_ZIP_ENTRY_BYTES || total > MAX_ZIP_TOTAL_BYTES) throw new Error('Not a valid xlsx file: the uncompressed content is too large');
      return false;
    },
  });
  return unzipSync(bytes);
}
