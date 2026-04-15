/**
 * Proxies Google Sheet CSV exports for the browser (avoids CORS).
 * Only allows Google Docs / Drive spreadsheet URLs.
 */
exports.handler = async (event) => {
  if (event.httpMethod !== "GET") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const raw = event.queryStringParameters?.url || "";
  let target;
  try {
    target = decodeURIComponent(raw);
  } catch {
    return { statusCode: 400, body: "Bad url encoding" };
  }

  if (!target || !/^https:\/\/(docs\.google\.com|drive\.google\.com)\//i.test(target)) {
    return { statusCode: 400, body: "Only Google Sheets URLs are allowed" };
  }

  try {
    const res = await fetch(target, { redirect: "follow" });
    const text = await res.text();
    if (!res.ok) {
      return { statusCode: res.status, body: text.slice(0, 8000) };
    }
    return {
      statusCode: 200,
      headers: {
        "Content-Type": "text/plain; charset=utf-8",
        "Cache-Control": "private, max-age=60",
      },
      body: text,
    };
  } catch (err) {
    return {
      statusCode: 502,
      body: String(err?.message || err),
    };
  }
};
