import { getStore } from "@netlify/blobs";

export default async (req) => {
  const store = getStore("journal-files");
  if (req.method === "GET") {
    const data = await store.get("template.xlsx", { type: "arrayBuffer" });
    if (!data) return new Response("missing", { status: 404 });
    return new Response(data, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Cache-Control": "no-store",
      },
    });
  }
  if (req.method === "PUT" || req.method === "POST") {
    const secret = process.env.TEMPLATE_ADMIN_KEY || "";
    const given = req.headers.get("x-admin-key") || "";
    if (!secret || given !== secret) {
      return new Response("Невірний пароль.", { status: 403 });
    }
    const buf = Buffer.from(await req.arrayBuffer());
    if (buf.length < 800) {
      return new Response("Файл замалий.", { status: 400 });
    }
    await store.set("template.xlsx", buf);
    return new Response("ok");
  }
  return new Response("method", { status: 405 });
};

export const config = { path: "/api/template" };
