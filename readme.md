# Simple Word Add-in

This is a simple Microsoft Word Add-in with three buttons:
1. Fetch data (token + URL)
2. Display fetched data
3. Insert HTML into Word document

## 🛠 Setup

1. Run a local HTTPS server:
```bash
npx http-server -S -C cert.pem -K key.pem

