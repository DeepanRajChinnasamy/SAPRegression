import shutil
import os

base_dir = "yaml-ui-app"
backend_dir = os.path.join(base_dir, "backend")
frontend_dir = os.path.join(base_dir, "frontend", "src")

os.makedirs(backend_dir, exist_ok=True)
os.makedirs(frontend_dir, exist_ok=True)

files = {
    os.path.join(backend_dir, "server.js"): """
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
""",

    os.path.join(backend_dir, "template.yaml"): """
name: "{{name}}"
age: "{{age}}"
email: "{{email}}"
address:
  city: "{{city}}"
  zip: "{{zip}}"
""",

    os.path.join(backend_dir, "package.json"): """
{
  "name": "yaml-ui-backend",
  "version": "1.0.0",
  "main": "server.js",
  "scripts": { "start": "node server.js" },
  "dependencies": {
    "axios": "^1.0.0",
    "body-parser": "^1.20.0",
    "cors": "^2.8.5",
    "express": "^4.18.2",
    "js-yaml": "^4.1.0"
  }
}
""",

    os.path.join(frontend_dir, "App.js"): """
import React, { useState } from 'react';
import axios from 'axios';

function App() {
  const [data, setData] = useState({ name:'', age:'', email:'', city:'', zip:'' });
  const [response, setResponse] = useState(null);

  const handleChange = e => setData({ ...data, [e.target.name]: e.target.value });
  const handleSubmit = async () => {
    try {
      const res = await axios.post('http://localhost:5000/send-request', data);
      setResponse(res.data);
    } catch (err) {
      setResponse({ error: err.message });
    }
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Send YAML‑based Request</h2>
      {Object.keys(data).map(key => (
        <div key={key} style={{ marginBottom: 10 }}>
          <label style={{ marginRight: 10 }}>{key}:</label>
          <input name={key} value={data[key]} onChange={handleChange} />
        </div>
      ))}
      <button onClick={handleSubmit}>Submit</button>
      <h3>Response:</h3>
      <pre>{JSON.stringify(response, null, 2)}</pre>
    </div>
  );
}

export default App;
""",

    os.path.join(frontend_dir, "index.js"): """
import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
""",

    os.path.join(base_dir, "frontend", "package.json"): """
{
  "name": "frontend",
  "version": "0.1.0",
  "private": true,
  "dependencies": {
    "axios": "^1.0.0",
    "react": "^18.0.0",
    "react-dom": "^18.0.0",
    "react-scripts": "5.0.1"
  },
  "scripts": {
    "start": "react-scripts start"
  }
}
"""
}

# Write each file
for path, content in files.items():
    with open(path, "w") as f:
        f.write(content.strip())

# Create zip
zip_file = shutil.make_archive("yaml-ui-app", 'zip', base_dir)
print(f"ZIP file created at: {zip_file}")
