FROM python:3.9-slim

WORKDIR /app

# Instalar dependências do sistema e curl para healthcheck
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements primeiro para melhor cache
COPY requirements.txt .

# Instalar dependências Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o restante da aplicação
COPY . .

# Criar diretórios necessários
RUN mkdir -p uploads static templates logs

# Expor a porta
EXPOSE 9090

# Comando para rodar a aplicação
CMD ["python", "-u", "app.py"]