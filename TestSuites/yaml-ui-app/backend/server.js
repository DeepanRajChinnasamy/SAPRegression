const express = require('express');
const fs = require('fs');
const yaml = require('js-yaml');
const axios = require('axios');
const cors = require('cors');
const bodyParser = require('body-parser');

const app = express();
app.use(cors());
app.use(bodyParser.json());

app.post('/send-request', async (req, res) => {
  try {
    let template = fs.readFileSync('template.yaml', 'utf8');
    Object.entries(req.body).forEach(([k, v]) => {
      template = template.replace(new RegExp(`{{${k}}}`, 'g'), v);
    });
    const obj = yaml.load(template);
    const response = await axios.post('https://httpbin.org/post', obj);
    res.json(response.data);
  } catch (err) {
    res.status(500).json({ error: err.toString() });
  }
});

app.listen(5000, () => console.log('Backend listening on http://localhost:5000'));