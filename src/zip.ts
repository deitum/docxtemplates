import JSZip from 'jszip';
import { type ZipInput } from './types';

const zipLoad = (inputFile: ZipInput) => JSZip.loadAsync(inputFile);
const zipGetText = (zip: JSZip, filename: string) => {
  const file_in_zip = zip.file(filename);
  if (!file_in_zip) return null;
  return file_in_zip.async('text');
};

/**
 * Writes a part of the package. Every caller hands over bytes that are already
 * UTF-8 encoded (`buildXml` returns a `Buffer`), so the content is declared
 * binary: telling JSZip it is text would invite it to encode it a second time.
 */
const zipSetText = (zip: JSZip, filename: string, data: ZipInput) =>
  zip.file(filename, data, { binary: true });

const zipSave = (zip: JSZip, compressionLevel: number) =>
  zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

// ==========================================
// Public API
// ==========================================
export { zipLoad, zipGetText, zipSetText, zipSave };
