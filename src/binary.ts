/** JSON-safe wire representation for file bytes crossing the webview boundary. */
export function encodeBytes(bytes: Uint8Array): string {
  const parts: string[] = [];
  // Avoid exceeding the function argument limit for large workbooks.
  for (let offset = 0; offset < bytes.length; offset += 32_768) {
    parts.push(String.fromCharCode(...bytes.subarray(offset, offset + 32_768)));
  }
  return btoa(parts.join(''));
}

export function decodeBytes(encoded: string): Uint8Array {
  if (typeof encoded !== 'string') throw new Error('Invalid spreadsheet data transfer. Reload the Sheets editor after updating the extension.');
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
