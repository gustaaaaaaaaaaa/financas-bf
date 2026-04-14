import express from "express";
import fetch from "node-fetch";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.json());
app.use(express.static("public"));

let accessToken = "";
let workbookSession = "";

// LOGIN
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

  const tokenData = await tokenRes.json();
  accessToken = tokenData.access_token;

  // CRIAR SESSION DE WORKBOOK (OBRIGATÓRIO)
  const sessionRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${process.env.FILE_ID}/workbook/createSession`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ persistChanges: true })
    }
  );

  const sessionData = await sessionRes.json();
  workbookSession = sessionData.id;

  res.redirect("/");
});

// ATUALIZAR EXCEL
app.post("/atualizar", async (req, res) => {
  if (!accessToken || !workbookSession) {
    return res.status(401).send({ error: "Sessão inválida" });
  }

  const { tipo, mes, valor } = req.body;

  const linha = 131 + Number(tipo);
  const coluna = String.fromCharCode(72 + Number(mes));
  const endereco = `${coluna}${linha}`;

  const url =
    `https://graph.microsoft.com/v1.0/me/drive/items/${process.env.FILE_ID}` +
    `/workbook/worksheets('Abril')/range(address='${endereco}')`;

  // LER VALOR
  const atualRes = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "workbook-session-id": workbookSession
    }
  });

  if (!atualRes.ok) {
    const err = await atualRes.text();
    return res.status(500).send(err);
  }

  const atual = await atualRes.json();
  const atualValor = atual.values?.[0]?.[0] || 0;
  const novoValor = Number(atualValor) + Number(valor);

  // ESCREVER
  const patchRes = await fetch(url, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      "workbook-session-id": workbookSession
    },
    body: JSON.stringify({ values: [[novoValor]] })
  });

  if (!patchRes.ok) {
    const err = await patchRes.text();
    return res.status(500).send(err);
  }

  res.send({ ok: true, novoValor });
});

app.listen(process.env.PORT || 3000);
