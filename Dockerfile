# --- Stage 1: build the webpack frontend ---
FROM node:20-slim AS frontend
WORKDIR /build
COPY package*.json ./
RUN npm ci
COPY webpack.config.js ./
COPY src/ ./src/
COPY assets/ ./assets/
COPY manifest.xml ./
RUN npm run build

# --- Stage 2: Python kernel server ---
FROM python:3.11-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server/ ./server/
# Copy built frontend so FastAPI can serve it as static files
COPY --from=frontend /build/dist ./dist

EXPOSE 8008

CMD ["uvicorn", "server.main:app", "--host", "0.0.0.0", "--port", "8008"]
