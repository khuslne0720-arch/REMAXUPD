# 🏠 Mongolian Real Estate Contract Generator

AI-powered web app for generating Mongolian real estate brokerage contracts.

## Features
- Upload a property certificate image → AI extracts structured data
- Edit/fill in any missing fields
- Choose contract type (Sell / Rent) and tier (Standard / Exclusive)
- Download a ready-to-sign `.docx` Word document
- 4 contract templates: sell_standard, sell_exclusive, rent_standard, rent_exclusive

## Requirements
- Node.js 18+
- An Anthropic API key

## Setup

```bash
# 1. Install dependencies
npm install

# 2. Set your Anthropic API key
export ANTHROPIC_API_KEY=sk-ant-...

# 3. Start the server
npm start
```

Then open http://localhost:3000 in your browser.

## Environment Variables

| Variable | Description |
|---|---|
| `ANTHROPIC_API_KEY` | Your Anthropic API key (required) |
| `PORT` | Server port (default: 3000) |

## Project Structure

```
├── server.js          # Express server, /analyze and /generate endpoints
├── templates.js       # All 4 Mongolian contract templates + fillTemplate()
├── docxGenerator.js   # Converts filled template → .docx using docx library
├── public/
│   └── index.html     # Full frontend (Vanilla JS, no framework)
└── uploads/           # Temp folder for image uploads (auto-cleaned)
```

## Contract Templates

Each template uses these placeholders:

| Placeholder | Source | Description |
|---|---|---|
| `{{name}}` | AI / user | Owner full name |
| `{{register}}` | AI / user | Registration ID |
| `{{address}}` | AI / user | Property address |
| `{{area}}` | AI / user | Area in m² |
| `{{cert}}` | AI / user | Certificate number |
| `{{phone}}` | User | Phone number |
| `{{email}}` | User | Email address |
| `{{price}}` | User | Price in MNT |
| `{{date}}` | Auto | Today's date |

## Adding Custom Templates

Edit `templates.js` — add a new key to `TEMPLATES` following the same structure.
Each section can have: `heading`, `body` (newline-separated), `fields` (array), `signature` (boolean).
