import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.json());
app.use(express.static("public"));

/* =========================
   LOGIN MICROSOFT
   ========================= */
app.get("/login", (req, res) => {
  const url =
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize" +
    `?client_id=${process.env.CLIENT_ID}` +
    "&response_type=code" +
    `&redirect_uri=${process.env.REDIRECT_URI}` +
    "&response_mode=query" +
    "&scope=Files.ReadWrite";

  res.redirect(url);
});

/* =========================
   CALLBACK – LOGIN + GRAVA EXCEL
   ========================= */
app.get("/callback", async (req, res) => {
  const { code, tipo, mes, valor } = req.query;

  if (!code || tipo === undefined || mes === undefined || valor === undefined) {
    return res.send("Parâmetros ausentes");
  }

  // 1️⃣ Troca code por access_token
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

  const tokenData = await tokenRes.json();
  const accessToken = tokenData.access_token;

  if (!accessToken) {
    return res.send("Erro ao obter access token");
  }

  // 2️⃣ Calcula célula
  const linha = 131 + Number(tipo);
  const coluna = String.fromCharCode(72 + Number(mes));
  const endereco = `${coluna}${linha}`;

  const url =
    `https://graph.microsoft.com/v1.0/me/drive/items/${process.env.FILE_ID}` +
    `/workbook/worksheets('Abril')/range(address='${endereco}')`;

  // 3️⃣ Lê valor atual
  const atualRes = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  const atual = await atualRes.json();
  const atualValor = atual.values?.[0]?.[0] || 0;
  const novoValor = Number(atualValor) + Number(valor);

  // 4️⃣ Grava soma
  await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: [[novoValor]] })
  });

  res.redirect("/?sucesso=1");
});

/* =========================
   SERVER
   ========================= */
app.listen(process.env.PORT || 3000);
``
