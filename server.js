const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');

const app = express();
app.use(cors());

const EXCEL_URL = 'https://gplaneta-my.sharepoint.com/:x:/g/personal/uveu1a_psoplaneta_com/Ef9O8SDtriBFodzJXjGq5BEBqZ1gYMZiqHsExnyqYnjz2A';

app.get('/excel', async (req, res) => {
  try {
    const response = await fetch(EXCEL_URL);
    if (!response.ok) {
      return res.status(response.status).send('Error al obtener el Excel');
    }

    const buffer = await response.arrayBuffer();
    res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buffer));
  } catch (error) {
    console.error('Error fetching Excel:', error);
    res.status(500).send('Error interno del servidor');
  }
});

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Servidor proxy corriendo en http://localhost:${PORT}`);
});