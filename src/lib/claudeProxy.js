// Calls the Anthropic API through our own Netlify serverless proxy (netlify/functions/claude-proxy.js)
// instead of api.anthropic.com directly — the API key lives only on the server, never in this bundle.
export async function callClaude(messagesPayload) {
  const response = await fetch("/.netlify/functions/claude-proxy", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(messagesPayload),
  });

  const raw = await response.text();
  let data;
  try {
    data = JSON.parse(raw);
  } catch (err) {
    // A plain Netlify Drop deploy (no GitHub/Functions) has no /.netlify/functions/* route at
    // all, so this request falls through to the SPA's own index.html (status 200, HTML body) —
    // that's the exact "Unexpected token '<'" failure. Give a clear, actionable message instead.
    throw new Error(
      "الميزة الذكية غير متاحة على هذا النشر. هذا يحدث عندما يُنشر الموقع بطريقة السحب المباشر لملف dist بدون GitHub — تلك الطريقة لا تشغّل دوال الخادم. اتبع خطوات GitHub + Netlify Functions + مفتاح ANTHROPIC_API_KEY الموضحة في README لتفعيل هذه الميزة."
    );
  }

  if (!response.ok) {
    throw new Error(data?.error?.message || data?.error || "تعذر الاتصال بخدمة الذكاء الاصطناعي");
  }
  return data;
}

export function extractText(data) {
  return (data.content || [])
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("");
}

export function parseJsonResponse(data) {
  const text = extractText(data);
  const clean = text.replace(/```json|```/g, "").trim();
  return JSON.parse(clean);
}

export function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}
