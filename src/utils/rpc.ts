// utils/rpc.ts
export async function fileToBase64(file: File): Promise<string> {
  const buf = await file.arrayBuffer()
  // 빠른 base64 인코딩
  let binary = ""
  const bytes = new Uint8Array(buf)
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}
