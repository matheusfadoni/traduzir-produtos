FROM python:3.10-slim

# Instala dependências
RUN apt-get update && apt-get install -y build-essential cron && apt-get clean

# Define timezone
ENV TZ=America/Asuncion

# Cria diretório
WORKDIR /app

# Copia tudo
COPY . /app

# Instala libs Python
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copia e ativa crontab
RUN chmod 0644 /app/crontab.txt && crontab /app/crontab.txt

# Inicia cron em foreground
CMD ["cron", "-f"]

