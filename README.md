# Unlock Your CV

Tailors your CV to any job description using AI. 3 rounds of targeted questions, then a single-page ATS-optimized CV.

## Local Development

```bash
pip install -r requirements.txt
pip install docx2pdf   # Windows only — for accurate page counting
python app.py
```

Open http://localhost:5050

## Deploy to Railway

1. Push this repo to GitHub
2. Connect the repo at [railway.app](https://railway.app)
3. Set environment variables in Railway dashboard:
   - `ANTHROPIC_API_KEY`
   - `STRIPE_SECRET_KEY`
   - `STRIPE_PUBLISHABLE_KEY`
   - `STRIPE_PRODUCT_ID`
4. Railway auto-detects Python, installs LibreOffice via `nixpacks.toml`, and runs via `Procfile`
5. Add your custom domain in Railway settings

## Architecture

- **Sonnet** handles Q&A rounds (fast, cheap)
- **Opus** handles final CV generation (highest quality)
- **LibreOffice** converts .docx → .pdf for page count verification (server)
- **docx2pdf** uses MS Word for page count (Windows local dev)
- **Stripe Checkout** handles payments (€0.99 per CV)

## Project Structure

```
├── app.py                  # Flask backend + Claude API + Stripe
├── cv_generator.py         # python-docx CV generation
├── templates/
│   ├── index.html          # Main UI
│   └── payment_result.html # Post-payment page
├── Procfile                # gunicorn for production
├── nixpacks.toml           # LibreOffice install for Railway
├── runtime.txt             # Python version
├── requirements.txt        # Python dependencies
└── .gitignore
```
