FROM python:3.9-slim

WORKDIR /app

# Instalar dependências do sistema incluindo libs para Excel
RUN apt-get update && apt-get install -y \
    curl \
    libgomp1 \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Copiar requirements primeiro para cache
COPY requirements.txt .

# Instalar dependências Python
RUN pip install --no-cache-dir --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o código da aplicação
COPY . .

# Criar diretórios necessários
RUN mkdir -p /tmp/uploads /tmp/logs

# Dar permissões
RUN chmod -R 755 /tmp/uploads /tmp/logs

# Expor a porta
EXPOSE 9090

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD curl -f http://localhost:9090/health || exit 1

# Comando para rodar a aplicação
CMD ["python", "app.py"]
