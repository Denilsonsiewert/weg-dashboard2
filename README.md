# WEG - Sistema de Controle Analítico Estanhagem (Versão Cloud)

Sistema completo para exportar dados do Excel para a nuvem e visualizar em dashboard web de qualquer lugar.

## 🌟 Arquitetura da Solução

```
Excel (VBA) → HTTP POST → API Cloud → Navegador Web
```

Agora você pode:
- ✅ Abrir a planilha em **qualquer computador**
- ✅ Dados enviados automaticamente para a **nuvem**
- ✅ Dashboard acessível de **qualquer navegador**
- ✅ **Sem dependência** de servidor local

## 📁 Arquivos do Projeto

```
├── server.js                  # Backend Node.js (API)
├── package.json              # Dependências do Node.js
├── public/
│   └── index.html           # Dashboard WEG (frontend)
├── MACRO_VBA_ATUALIZADA.vba # Macro para colar no Excel
└── README.md                # Este arquivo
```

## 🚀 Como Fazer o Deploy (Hospedagem Gratuita)

### Opção 1: Render.com (Recomendado - Gratuito)

1. **Criar conta no Render**: https://render.com
2. **New > Web Service**
3. **Conectar este repositório Git** (ou fazer upload dos arquivos)
4. **Configurações**:
   - Name: `weg-dashboard`
   - Environment: `Node`
   - Build Command: `npm install`
   - Start Command: `npm start`
   - Plan: `Free`
5. **Deploy**
6. **Copiar a URL** gerada (ex: `https://weg-dashboard.onrender.com`)

### Opção 2: Railway.app (Gratuito)

1. **Criar conta**: https://railway.app
2. **New Project > Deploy from GitHub**
3. **Selecionar repositório**
4. **Deploy automático**
5. **Copiar a URL** gerada

### Opção 3: Heroku (Gratuito com limitações)

1. **Criar conta**: https://heroku.com
2. **New > Create new app**
3. **Deploy via GitHub ou CLI**
4. **Copiar a URL** gerada

## 📝 Configuração da Macro VBA

Depois de fazer o deploy:

1. **Abra o Excel**
2. **Pressione** `Alt + F11` (abre o editor VBA)
3. **Insert > Module**
4. **Cole** o conteúdo de `MACRO_VBA_ATUALIZADA.vba`
5. **IMPORTANTE**: Na linha 7, altere a URL:

```vba
Const URL_API As String = "https://SEU-SERVIDOR.com/api/dados"
```

Substitua `SEU-SERVIDOR.com` pela URL que você recebeu do Render/Railway/Heroku.

**Exemplo**:
```vba
Const URL_API As String = "https://weg-dashboard.onrender.com/api/dados"
```

6. **Salve** a planilha como `.xlsm` (Excel com macros)

## ▶️ Como Usar

### No Excel:

1. **Preencha** os dados normalmente nas planilhas
2. **Execute** a macro: `Alt + F8` > `ExportarDadosParaAPI` > `Executar`
3. **Aguarde** a mensagem de confirmação: "✅ Dados enviados com sucesso!"

### No Navegador:

1. **Acesse** a URL do seu servidor (ex: `https://weg-dashboard.onrender.com`)
2. **Visualize** o dashboard atualizado
3. **Os dados** são atualizados automaticamente a cada 60 segundos
4. **Slides** trocam automaticamente a cada 20 segundos

## 🔧 Desenvolvimento Local (Teste)

Para testar localmente antes do deploy:

```bash
# Instalar dependências
npm install

# Iniciar servidor
npm start
```

Acesse: http://localhost:3000

Na macro VBA, use:
```vba
Const URL_API As String = "http://localhost:3000/api/dados"
```

## 📊 Endpoints da API

- **POST** `/api/dados` - Recebe dados do Excel (VBA)
- **GET** `/api/dados` - Retorna os dados (frontend)
- **GET** `/api/health` - Status do servidor

## 🔒 Segurança

### Produção (Recomendações):

Para ambiente de produção, considere adicionar:

1. **Autenticação**: Token API ou Basic Auth
2. **HTTPS**: Obrigatório (Render/Railway já fornecem)
3. **Rate Limiting**: Limitar requisições por IP
4. **Validação**: Validar estrutura do JSON

### Exemplo com Token (Opcional):

No `server.js`, adicione antes do endpoint:

```javascript
const API_TOKEN = process.env.API_TOKEN || 'seu-token-secreto-aqui';

app.use('/api/dados', (req, res, next) => {
    const token = req.headers['authorization'];
    if (token !== `Bearer ${API_TOKEN}`) {
        return res.status(401).json({ error: 'Não autorizado' });
    }
    next();
});
```

Na macro VBA:

```vba
http.setRequestHeader "Authorization", "Bearer seu-token-secreto-aqui"
```

## 🆘 Solução de Problemas

### Erro: "Não foi possível enviar dados"

1. **Verifique** se a URL está correta na macro
2. **Teste** a API no navegador: `https://seu-servidor.com/api/health`
3. **Confirme** que o servidor está ativo (Render/Railway podem adormecer após inatividade)

### Erro: "MSXML2.XMLHTTP.6.0 não encontrado"

- Tente usar: `CreateObject("MSXML2.ServerXMLHTTP")`
- Ou: `CreateObject("WinHttp.WinHttpRequest.5.1")`

### Dashboard não atualiza:

1. **Limpe** o cache do navegador (Ctrl + F5)
2. **Verifique** se há dados no servidor: `/api/dados`
3. **Abra** o console (F12) e veja os erros

## 📱 Compatibilidade

- ✅ Excel 2010 ou superior
- ✅ Todos os navegadores modernos (Chrome, Firefox, Edge, Safari)
- ✅ Mobile (responsivo)

## 📄 Licença

MIT - Uso livre

## 🤝 Suporte

Dúvidas? Entre em contato com o time de TI da WEG.
