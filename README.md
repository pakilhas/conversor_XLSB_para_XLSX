#  Conversor XLSB para XLSX

> **Documentação Oficial — Atualiza**

---

---

##  Visão Geral

O **Conversor XLSB para XLSX** é uma aplicação web desenvolvida em **Python/Flask**, projetada para converter arquivos Excel no formato **binário (.xlsb)** para o formato **XML (.xlsx)**.  

A aplicação é **containerizada com Docker** e disponibiliza uma **interface web moderna**, permitindo o **acompanhamento em tempo real** da conversão.

---

##  Funcionalidades

###  Conversão de Arquivos
- **XLSB → XLSX:** Conversão precisa de arquivos binários para XML.  
- **Múltiplas Planilhas:** Preserva todas as abas do arquivo original.  
- **Integridade de Dados:** Mantém a estrutura e os dados intactos.  

###  Interface Web
- **Upload Drag & Drop:** Arraste e solte seus arquivos facilmente.  
- **Barra de Progresso em Tempo Real:** Visualize o avanço da conversão.  
- **Tempo de Processamento:** Mostra o tempo total da conversão.  
- **Feedback Visual:** Notificações automáticas de sucesso ou erro.  
- **Design Responsivo:** Compatível com desktop e dispositivos móveis.  

###  Gerenciamento de Container
- **Docker Compose:** Facilita a orquestração dos containers.  
- **Health Checks:** Verificação automática de integridade.  
- **Restart Automático:** Reinício em caso de falhas.  
- **Logs Centralizados:** Armazenamento persistente e monitorado.  
- **Persistência de Dados:** Mantém uploads e logs salvos.  

---

##  Tecnologias Utilizadas

### Backend
- **Python 3.9** — Linguagem principal  
- **Flask 2.3.3** — Framework web  
- **Pandas 1.5.3** — Manipulação de dados  
- **PyXLSB 1.0.10** — Leitura de arquivos .xlsb  
- **OpenPyXL 3.1.2** — Escrita de arquivos .xlsx  

### Frontend
- **HTML5 / CSS3 / JavaScript** — Estrutura e interatividade  
- **CSS Grid / Flexbox** — Layout responsivo  

### Infraestrutura
- **Docker** — Containerização  
- **Docker Compose** — Orquestração  
- **Python Slim** — Imagem base otimizada  

---

##  Pré-requisitos

### Software Necessário
- **Docker Desktop** (Windows/Mac) ou **Docker Engine** (Linux)  
- **Git** (opcional, para clonar o repositório)

### Recursos do Sistema
- **RAM:** 2GB mínimo (4GB recomendado)  
- **Armazenamento:** 1GB livre  
- **Rede:** Porta **9090** liberada  

---

##  Instalação e Execução

### Método 1: Docker Compose (Recomendado)
```bash
# 1. Clone o repositório
git clone <url-do-repositorio>
cd meu_conversor

# 2. Inicie o container
docker-compose up -d

# 3. Verifique se está rodando
docker-compose ps
