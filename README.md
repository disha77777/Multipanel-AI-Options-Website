# Multipanel AI

Multipanel AI is a Streamlit workspace for AI-assisted research, Excel reports, presentation generation, image-table OCR, and data cleaning. It previews generated artifacts in the browser before download.

## Features

- Generate a PowerPoint deck from a topic or thesis statement.
- Research a question and export a styled Excel workbook with data, sources, and charts.
- Extract tables from images into Excel with Gemini OCR.
- Create a research plan or a source-backed verified research report.
- Upload CSV or Excel data and apply natural-language cleaning operations.

## Setup

Use Python 3.12, then install the dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Create a `.env` file from [`.env.example`](.env.example). Gemini handles live, source-grounded research and Excel research. Add an xAI key to enable the provider selector for presentations, research plans, data cleaning, and image/table reading:

```env
GEMINI_API_KEY=your_gemini_api_key
XAI_API_KEY=your_xai_api_key
```

Run the app:

```bash
streamlit run app.py
```

Open the URL printed by Streamlit, usually `http://localhost:8501`.

## React UI with the Python API

The React interface in [`AI Multipanel for Document Creation`](AI%20Multipanel%20for%20Document%20Creation) uses `api_server.py`. Run both services during development:

```bash
python3 -m pip install -r requirements.txt
uvicorn api_server:app --reload --port 8000
```

In a second terminal:

```bash
cd "AI Multipanel for Document Creation"
npm install
npm run dev
```

Vite proxies `/api` to `http://127.0.0.1:8000` by default. Set `API_PROXY_TARGET` if the API runs elsewhere. For separately hosted production UI and API, set `VITE_API_BASE_URL` to the API origin when building the UI and set `ALLOWED_ORIGINS` on the API to the exact deployed UI origin(s), comma-separated. The UI checks `/api/health` and reports whether its backend is reachable.

## Deploy the React UI and API to Render

The root [`render.yaml`](render.yaml) Blueprint defines a free static UI and a free Python API. It derives `VITE_API_BASE_URL` and the API's exact `ALLOWED_ORIGINS` from the two Render service URLs. The API health check is `/api/health`.

1. Push the branch containing `render.yaml` and the app code to a GitHub repository connected to Render.
2. In Render, choose **New > Blueprint**, select that repository and branch, and review the two services.
3. Enter `GEMINI_API_KEY` and/or `XAI_API_KEY` in Render's secret prompt. At least one provider key is needed for AI generation; never put real keys in this file or commit them.
4. Deploy, then confirm the API service's `/api/health` reports `status: ok` and the configured provider(s).

The free API service can sleep when idle, and its in-memory sessions and generated files are not durable. This setup is suitable for a low-cost demo, not persistent multi-user data. If you add a custom UI domain, update the API's `ALLOWED_ORIGINS` to include that exact `https://` origin.

## Release deployment

### Streamlit app

Deploy this repository using Streamlit Community Cloud with `app.py` as the main file. Add `GEMINI_API_KEY` under **Settings > Secrets** using TOML syntax:

```toml
GEMINI_API_KEY = "your_gemini_api_key"
```

Do not commit `.env`, `key.env`, or Streamlit secrets.

### Chemistry API

The optional chemistry chat frontend/API is deployed separately through the existing Vercel configuration. Set `GEMINI_API_KEY` in the Vercel project environment, then deploy from the repository root. Verify the deployment with `/api/health` before sharing it.

The API limits uploads to 8 MB, accepts JPEG/PNG/WebP images, rate-limits requests per instance, and returns generic errors without exposing model or server details.

## Release checks

```bash
python3 -m compileall -q app.py api chem.py datacleaner.py excel.py presentation.py reader.py research_assistant.py research_planner.py pinkman/pinkman_server.py
python3 -m pip install -r requirements.txt
streamlit run app.py
```

The `.env` file is ignored by Git and must not be committed.

## Security baseline

This MVP implements the application-layer controls below:

1. **Secrets and rotation:** only `GEMINI_API_KEY` is read by AI features; keep it in Streamlit/Vercel secrets, never in source control, and rotate it on a schedule or immediately after suspected exposure. Use separate keys for development and production.
2. **Private networking:** the app has no database. Keep any future database and private services in private subnets with no public ingress. The Gemini API is an intentional external dependency, so “no internet between services” cannot apply to model calls without adopting a private model endpoint.
3. **WAF:** place Vercel WAF, Cloudflare, or your cloud edge firewall in front of the API. Enforce managed OWASP rules, bot protection, request-size limits, and country/IP policy at the edge. A response header cannot replace a WAF.
4. **Rate limiting:** the chemistry API uses configurable per-IP limits. For multiple serverless instances, add a distributed limiter at the edge (for example, a provider-native rate-limit product) because in-memory counters are instance-local.
5. **Input validation:** question length, image MIME type, image decoding, image dimensions, structured mechanism output, source URLs, citations, and data-cleaning operations are validated before use.
6. **CORS:** set `ALLOWED_ORIGINS` to exact origins such as `https://your-app.example`; do not use `*`. The API rejects unconfigured browser origins.
7. **Configuration and flags:** use environment variables for limits and `ENABLE_CHEMISTRY` / `ENABLE_LIVE_RESEARCH` for emergency feature shutdowns. Keep production and development environments separate.
8. **Prompt injection:** model prompts delimit user content as untrusted data, reject unsupported structured operations, validate generated JSON, and instruct models not to reveal system prompts or secrets. This reduces risk but cannot guarantee model immunity.

### Required edge policy before public launch

- WAF enabled in blocking mode with OWASP managed rules.
- TLS-only custom domain; no public admin or database endpoint.
- Distributed rate limits on `/api/chemistry`, `/api/chemistry/stream`, and `/api/mechanism/generate`.
- Request body limit of 8 MB at the edge, matching the application.
- Alerting for 4xx/5xx spikes, rate-limit events, and provider spend.
- A tested key-revocation and rotation runbook.
