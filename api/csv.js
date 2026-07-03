// Vercel Serverless Function: đọc CSV từ Google Sheet phía server (không dính CORS)
// Gọi: /api/csv?gid=<gid>
export default async function handler(req, res) {
  const KEY = "2PACX-1vSve6XRHg5gWRzqkazHm5zvlrkTkAMLa7TJms_U-ebAFcrDAmcvCYfNJ50hrvV988tXyKC7q70LQgPc";
  // Chỉ cho phép các tab đã biết (an toàn)
  const ALLOW = new Set([
    "460836856",   // Data theo SPDV
    "1758921427",  // Data theo NCC
    "163849763",   // Bảng tỉ giá
    "562469906",   // Data Robux
    "61864847",    // Data Nick Roblox
    "1043029815",  // Data giftcard
    "1868031300",  // Data Razergold
    "793401472",   // Data OGGAMING
    "1289659560",  // Data RBX Daily
    "1711960798"   // Logic
  ]);
  const gid = String((req.query && req.query.gid) || "");
  if (!ALLOW.has(gid)) {
    res.status(400).send("gid khong hop le");
    return;
  }
  const url = `https://docs.google.com/spreadsheets/d/e/${KEY}/pub?gid=${gid}&single=true&output=csv`;
  try {
    const r = await fetch(url, { redirect: "follow" });
    const text = await r.text();
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader("Cache-Control", "s-maxage=30, stale-while-revalidate=120");
    res.status(200).send(text);
  } catch (e) {
    res.status(502).send("Loi doc Google Sheet");
  }
}
