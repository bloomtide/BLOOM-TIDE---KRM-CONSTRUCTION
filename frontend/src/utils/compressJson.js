/**
 * Compress JSON with gzip and base64 for smaller payloads.
 * Uses browser CompressionStream when available; otherwise returns null (send uncompressed).
 */

const CHUNK_SIZE = 8192

function uint8ArrayToBase64(bytes) {
  let binary = ''
  for (let i = 0; i < bytes.length; i += CHUNK_SIZE) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK_SIZE))
  }
  return btoa(binary)
}

/**
 * Compress a JSON-serializable object to a gzip'd base64 string.
 * @param {object} obj - Object to compress (will be JSON.stringify'd)
 * @returns {Promise<string|null>} Base64-encoded gzip payload, or null if compression not supported
 */
export async function compressJsonToBase64(obj) {
  if (typeof CompressionStream === 'undefined') return null
  try {
    const jsonString = JSON.stringify(obj)
    const blob = new Blob([jsonString], { type: 'application/json' })
    const stream = blob.stream().pipeThrough(new CompressionStream('gzip'))
    const compressed = await new Response(stream).arrayBuffer()
    return uint8ArrayToBase64(new Uint8Array(compressed))
  } catch (e) {
    console.warn('[compressJson] Compression failed:', e)
    return null
  }
}
