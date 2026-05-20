
const express = require("express");
const https = require("https");
const selfsigned = require("selfsigned");
const path = require("path");

const app = express();

// Servir archivos del folder src
app.use(express.static(path.join(__dirname, "src")));

// Certificado HTTPS local
const attrs = [{ name: "commonName", value: "localhost" }];
const pems = selfsigned.generate(attrs, { days: 365 });

// Levantar servidor
https.createServer(
  { key: pems.private, cert: pems.cert },
  app
).listen(3000, () => {
  console.log("✅ Servidor HTTPS listo en https://localhost:3000");
});
