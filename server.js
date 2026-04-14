
import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";
import cookieParser from "cookie-parser";

dotenv.config();

const app = express();
app.use(express.json());
app.use(cookieParser());
app.use(express.static("public"));

/* =========================================================
   LOGIN MICROSOFT
   ========================================================= */

app.get("/login", (req, res) => {
  const url =
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize" +
    `?client_id=${process.env.CLIENT_ID}` +
    "&response_type=code" +
    `&redirect_uri=${process.env.REDIRECT_URI}` +
    "&response_mode=query" +
    "&scope=Files.ReadWrite offline_access" +
    "&prompt=consent";

  res.redirect(url);
});

/* =========================================================
   CALLBACK — troca code por refresh_token
   ========================================================= */
app.get("/callback", async (req, res) => {
  const code = req.query.code;

  const tokenRes = await fetch(
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:
        `client_id=${process.env.CLIENT_ID}` +
        `&client_secret=${process.env.CLIENT_SECRET}` +
        `&code=${code}` +
        `&redirect_uri=${process.env.REDIRECT_URI}` +
        "&grant_type=authorization_code"
    }
  );

  const data = await tokenRes.json();

  if (!data.refresh_token) {
    return res
      .status(500)
      .send("Erro ao obter refresh_token do Microsoft");
  }

  // Guarda refresh token em cookie seguro
  res.cookie("refresh_token", data.refresh_token, {
    httpOnly: true,
    secure: true,
    sameSite: "lax"
  });

  res.redirect("/");
});

/* =========================================================
   ATUALIZAR EXCEL (STATELESS, ROBUSTO)
   ========================================================= */
app.post("/atualizar", async (req, res) => {
  const refreshToken = req.cookies.refresh_token;

  if (!refreshToken) {
    return res.status(401).send("Usuário não autenticado");
  }

  // 🔄 gera access token novo usando refresh_token
  const tokenRes = await fetch(
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body:
        `client_id=${process.env.CLIENT_ID}` +
        `&client_secret=${process.env.CLIENT_SECRET}` +
        `&refresh_token=${refreshToken}` +
        "&grant_type=refresh_token"
    }
  );

  const tokenData = await tokenRes.json();
  const accessToken = tokenData.access_token;

  if (!accessToken) {
    return res.status(401).send("Token inválido");
  }

  const { tipo, mes, valor } = req.body;

  const linha = 131 + Number(tipo);
  const coluna = String.fromCharCode(72 + Number(mes)); // H = Março
  const endereco = `${coluna}${linha}`;

  const url =
    `https://graph.microsoft.com/v1.0/me/drive/items/${process.env.FILE_ID}` +
    `/workbook/worksheets('Abril')/range(address='${endereco}')`;

  // 🔍 Lê valor atual
  const atualRes = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  if (!atualRes.ok) {
    const err = await atualRes.text();
    return res.status(500).send(err);
  }

  const atual = await atualRes.json();
  const atualValor = atual.values?.[0]?.[0] || 0;
  const novoValor = Number(atualValor) + Number(valor);

  // ✏️ Grava soma
  const patchRes = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: [[novoValor]] })
  });

  if (!patchRes.ok) {
    const err = await patchRes.text();
    return res.status(500).send(err);
  }

  res.send({ ok: true, novoValor });
});

/* =========================================================
   SERVER
   ========================================================= */
app.listen(process.env.PORT || 3000);
