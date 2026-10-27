FROM node:22.12.0-bookworm-slim AS frontend-build

WORKDIR /frontend
RUN npm install --global pnpm@9.15.9
COPY ["AI Multipanel for Document Creation/package.json", "AI Multipanel for Document Creation/pnpm-lock.yaml", "./"]
RUN pnpm install --frozen-lockfile
COPY ["AI Multipanel for Document Creation/", "./"]
RUN pnpm run build

FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY api_server.py ai_provider.py datacleaner.py excel.py presentation.py presentation_media.py reader.py research_assistant.py research_planner.py security.py structured_output.py ./
COPY --from=frontend-build ["/frontend/dist", "/app/AI Multipanel for Document Creation/dist"]

EXPOSE 10000
CMD ["sh", "-c", "exec uvicorn api_server:app --host 0.0.0.0 --port ${PORT:-10000}"]
