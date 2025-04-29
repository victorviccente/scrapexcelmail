# 📊 scrapeexcelmail

Um projeto de Web Scraping que coleta dados da web, salva os resultados em uma planilha `.xlsx` e em um arquivo `.csv`, e envia os arquivos por e-mail automaticamente.

## 🚀 Funcionalidades

- Coleta de dados via Web Scraping
- Geração de arquivos Excel e CSV com os dados extraídos
- Envio automático dos arquivos por e-mail usando SMTP

## ⚙️ Requisitos

- Python 3.8+
- Conta de e-mail (recomendado Gmail com senha de app)

## 🔐 Configuração do `.env`

Antes de executar o projeto, crie um arquivo chamado `.env` na raiz do projeto com as seguintes variáveis:

```env
SENDER_EMAIL=seu_email@gmail.com
APP_PASSWORD=sua_senha_de_aplicativo
RECIPIENT_EMAIL=email_do_destinatario@gmail.com
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
