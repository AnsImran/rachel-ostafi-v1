FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    FLASK_APP=app \
    FLASK_RUN_HOST=0.0.0.0 \
    UV_SYSTEM_PYTHON=1

WORKDIR /app

COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000

CMD ["flask", "--app", "app", "run", "--debug", "--host", "0.0.0.0", "--port", "5000"]
