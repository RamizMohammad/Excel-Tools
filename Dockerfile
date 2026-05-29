FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server/ ./server/

EXPOSE 8008

# Run from /app so `import server.X` resolves correctly
CMD ["uvicorn", "server.main:app", "--host", "0.0.0.0", "--port", "8008"]
