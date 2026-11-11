/**
 * Proteção por senha do site do Grupo Transcender (Vercel Routing Middleware).
 * Nenhuma página do site é entregue sem a senha. A senha fica na variável de ambiente
 * SENHA_SITE do projeto na Vercel (Settings → Environment Variables). Trocar a senha lá
 * desconecta todos os aparelhos na próxima visita.
 */
export const config = { matcher: "/:path*" };

const COOKIE = "transcender_acesso";
const DIAS = 30;
const LIVRES = ["/assets/logo.png", "/assets/bg-porta.jpg", "/favicon.ico"];

async function tokenDa(senha) {
  const dados = new TextEncoder().encode("grupo-transcender:" + senha);
  const hash = await crypto.subtle.digest("SHA-256", dados);
  return [...new Uint8Array(hash)].map(b => b.toString(16).padStart(2, "0")).join("");
}

function lerCookie(req, nome) {
  const c = req.headers.get("cookie") || "";
  const m = c.match(new RegExp("(?:^|;\\s*)" + nome + "=([^;]+)"));
  return m ? decodeURIComponent(m[1]) : "";
}

export default async function middleware(req) {
  const url = new URL(req.url);
  if (LIVRES.includes(url.pathname)) return;

  const senha = process.env.SENHA_SITE;
  if (!senha) return new Response("Senha do site não configurada.", { status: 503 });
  const token = await tokenDa(senha);

  // envio do formulário de senha
  if (url.pathname === "/entrar" && req.method === "POST") {
    const form = await req.formData();
    const destino = String(form.get("destino") || "/");
    const volta = destino.startsWith("/") && !destino.startsWith("//") ? destino : "/";
    if (String(form.get("senha") || "").trim() === senha) {
      return new Response(null, {
        status: 303,
        headers: {
          "Location": volta,
          "Set-Cookie": `${COOKIE}=${token}; Path=/; Max-Age=${DIAS * 86400}; HttpOnly; Secure; SameSite=Lax`,
          "Cache-Control": "no-store"
        }
      });
    }
    return tela(volta, true);
  }

  if (lerCookie(req, COOKIE) === token) return; // acesso liberado: segue para o site
  return tela(url.pathname + url.search, false);
}

function tela(destino, erro) {
  const esc = s => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  const html = `<!DOCTYPE html>
<html lang="pt-BR"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex">
<title>Grupo Transcender de NA</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=Inter:wght@400;700;800&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{min-height:100vh;min-height:100dvh;display:grid;place-items:center;padding:24px;font-family:Inter,system-ui,sans-serif;color:#174A6B;
  background:radial-gradient(120% 90% at 10% 0%,#D9F3FE 0%,transparent 55%),radial-gradient(120% 90% at 90% 100%,#FCE1EC 0%,transparent 55%),#FFF8FB}
.card{width:100%;max-width:420px;background:rgba(255,255,255,.92);border-radius:28px;padding:34px 28px 28px;text-align:center;box-shadow:0 30px 70px -28px rgba(23,74,107,.45)}
.logo{width:120px;height:120px;margin:0 auto 14px;border-radius:50%;background:url(/assets/logo.png) center/cover no-repeat;box-shadow:0 14px 30px -14px rgba(209,79,138,.5)}
h1{font-family:"DM Serif Display",Georgia,serif;font-weight:400;font-size:38px;line-height:1.05;color:#1A78AE}
.sub{font-size:17px;font-weight:700;margin-top:6px}
form{margin-top:26px;display:flex;flex-direction:column;gap:12px;text-align:left}
label{font-size:14px;font-weight:700}
input{font:inherit;font-size:18px;color:#174A6B;padding:14px 16px;border-radius:14px;border:2px solid #D6E6F2;background:#fff;outline:none;width:100%}
input:focus{border-color:#5BCEFA}
button{font:inherit;font-size:18px;font-weight:800;color:#174A6B;background:#FAD3DE;border:2px solid #F5A9B8;border-radius:999px;padding:14px;cursor:pointer;margin-top:4px}
button:hover{background:#F7BFCF}
.erro{color:#D14F8A;font-weight:700;font-size:15px}
</style></head>
<body>
<main class="card">
  <div class="logo" role="img" aria-label="Borboleta do Grupo Transcender"></div>
  <h1>Grupo Transcender</h1>
  <p class="sub">de Narcóticos Anônimos</p>
  <form method="post" action="/entrar">
    <label for="senha">Senha de acesso</label>
    <input id="senha" name="senha" type="password" autocomplete="current-password" required autofocus>
    <input type="hidden" name="destino" value="${esc(destino)}">
    ${erro ? '<p class="erro">Senha incorreta. Tente de novo.</p>' : ""}
    <button type="submit">Entrar</button>
  </form>
</main>
<script>
  // guarda o #roteiro (a parte depois do # não chega ao servidor)
  if (location.hash) document.querySelector('input[name=destino]').value += location.hash;
</script>
</body></html>`;
  return new Response(html, { status: erro ? 401 : 200, headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" } });
}
