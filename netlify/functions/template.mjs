import { getStore } from "@netlify/blobs";

export default async (req, context) => {
  try {
    const store = getStore("journal-files");
    if (req.method === "GET") {
      const data = await store.get("template.xlsx", { type: "arrayBuffer" });
      if (!data) return new Response("missing", { status: 404 });
      const storedName =
        (await store.get("template-name", { type: "text" })) || "template.xlsx";
      return new Response(data, {
        headers: {
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Cache-Control": "no-store",
          "X-Template-Name": encodeURIComponent(storedName),
        },
      });
    }
    if (req.method === "DELETE") {
      await store.delete("template.xlsx");
      await store.delete("template-name");
      return new Response("ok");
    }
    if (req.method === "PUT" || req.method === "POST") {
      const buf = Buffer.from(await req.arrayBuffer());
      if (buf.length < 800) {
        return new Response("Файл замалий.", { status: 400 });
      }
      const rawName = req.headers.get("x-template-name") || "template.xlsx";
      let name = "template.xlsx";
      try {
        name = decodeURIComponent(rawName);
      } catch (e) {
        name = rawName;
      }
      name = String(name).replace(/^.*[/\\]/, "").trim() || "template.xlsx";
      if (!/\.xlsx?$/i.test(name)) name += ".xlsx";
      await store.set("template.xlsx", buf);
      await store.set("template-name", name);
      return new Response("ok");
    }
    return new Response("method", { status: 405 });
  } catch (err) {
    if (req.method === "GET") return new Response("missing", { status: 404 });
    return new Response(
      err && err.message ? err.message : "Помилка сховища шаблону.",
      { status: 500 }
    );
  }
};

export const config = { path: "/api/template" };
