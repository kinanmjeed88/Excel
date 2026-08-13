/**
 * يضيف رمزاً أو مرجع خلية إلى مسودة الصيغة مع ضمان بادئة = عند الحاجة.
 * تبقى الدالة مستقلة عن الواجهة لتكون تجربة أزرار الصيغ قابلة للاختبار.
 */
export function appendFormulaDraftToken(currentDraft: string, token: string) {
  let expression = currentDraft.trim();
  if (!expression) return token === "=" ? "=" : `=${token}`;
  if (!expression.startsWith("=")) expression = `=${expression}`;
  return token === "=" && expression === "=" ? expression : `${expression}${token}`;
}
