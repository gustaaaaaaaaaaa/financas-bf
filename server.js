import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.json());
app.use(express.static("public"));

let accessToken = "";

// LOGIN MICROSOFT
app.get("/login", (req, res) => {
  const url =
    "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize" +
    `?client_id=${process.env.CLIENT_ID}` +
    "&response_type=code" +
    `&redirect_uri=${process.env.REDIRECT_URI}` +
    "&response_mode=query" +
    "&scope=Files.ReadWrite offline_access";

  res.redirect(url);
});

// CALLBACK
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
  accessToken = data.access_token;

  res.redirect("/");
});

// ATUALIZAR EXCEL
app.post("/atualizar", async (req, res) => {
  const { tipo, mes, valor } = req.body;

  const linha = 131 + Number(tipo);
  const coluna = String.fromCharCode(72 + Number(mes));
  const endereco = `${coluna}${linha}`;

  const baseUrl =
    `https://graph.microsoft.com/v1.0/me/drive/items/${process.env.FILE_ID}` +
    `/workbook/worksheets('Abril')/range(address='${endereco}')`;

  // LER VALOR ATUAL
  const atualRes = await fetch(baseUrl, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  const atual = await atualRes.json();
  const atualValor = atual.values?.[0]?.[0] || 0;

  // SOMAR
  const novoValor = Number(atualValor) + Number(valor);

  // GRAVAR
  await fetch(baseUrl, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ values: [[novoValor]] })
  });

  res.send({ ok: true });
});

app.listen(process.env.PORT || 3000);
