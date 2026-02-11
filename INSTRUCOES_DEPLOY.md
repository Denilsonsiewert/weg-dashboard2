# 🚀 INSTRUÇÕES RÁPIDAS DE DEPLOY

## Passo a Passo Simplificado

### 1️⃣ Hospedar o Servidor (Escolha UMA opção)

#### **OPÇÃO A: Render.com** (Mais fácil - Recomendado)

1. Acesse: https://render.com
2. Crie uma conta gratuita
3. Clique em **"New +"** → **"Web Service"**
4. Conecte sua conta GitHub OU faça upload manual dos arquivos
5. Configure:
   - **Name**: `weg-dashboard` (ou outro nome)
   - **Environment**: `Node`
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Instance Type**: `Free`
6. Clique em **"Create Web Service"**
7. **Aguarde** o deploy (3-5 minutos)
8. **COPIE** a URL gerada (ex: `https://weg-dashboard.onrender.com`)

#### **OPÇÃO B: Railway.app**

1. Acesse: https://railway.app
2. **"Start a New Project"** → **"Deploy from GitHub repo"**
3. Conecte o repositório
4. Deploy automático
5. **COPIE** a URL gerada

#### **OPÇÃO C: Vercel** (Mais rápido)

1. Acesse: https://vercel.com
2. **"New Project"** → Importar repositório
3. Configure:
   - **Framework Preset**: Other
   - **Build Command**: `npm install`
   - **Output Directory**: `public`
4. Deploy automático
5. **COPIE** a URL gerada

---

### 2️⃣ Configurar a Macro VBA no Excel

1. **Abra** sua planilha Excel
2. Pressione **Alt + F11** (abre o editor VBA)
3. Menu **Insert** → **Module**
4. **Cole** todo o conteúdo do arquivo `MACRO_VBA_ATUALIZADA.vba`
5. **PROCURE** a linha 7:

```vba
Const URL_API As String = "https://seu-servidor.com/api/dados"
```

6. **SUBSTITUA** pela URL que você copiou no passo 1:

```vba
Const URL_API As String = "https://weg-dashboard.onrender.com/api/dados"
```

7. **Salve** a planilha como `.xlsm` (Excel Macro-Enabled Workbook)

---

### 3️⃣ Testar o Sistema

#### No Excel:

1. Pressione **Alt + F8**
2. Selecione **`ExportarDadosParaAPI`**
3. Clique em **Executar**
4. Deve aparecer: **"✅ Dados enviados com sucesso!"**

#### No Navegador:

1. Abra sua URL: `https://weg-dashboard.onrender.com`
2. O dashboard deve carregar com os dados do Excel
3. **Pronto!** ✨

---

## 🎯 Vantagens da Solução Cloud

| Antes (Local) | Depois (Cloud) |
|--------------|----------------|
| ❌ Só funciona no PC com servidor Python | ✅ Funciona em qualquer computador |
| ❌ Precisa iniciar servidor manualmente | ✅ Sempre disponível 24/7 |
| ❌ IP muda, precisa reconfigurar | ✅ URL fixa e permanente |
| ❌ Não funciona fora da rede local | ✅ Acesso de qualquer lugar (internet) |
| ❌ Depende do PC estar ligado | ✅ Servidor na nuvem sempre ativo |

---

## ❓ FAQ - Perguntas Frequentes

### **P: É realmente gratuito?**
**R:** Sim! Render, Railway e Vercel oferecem planos gratuitos suficientes para este projeto.

### **P: Os dados ficam salvos?**
**R:** Os dados ficam salvos na memória do servidor. Cada novo envio do Excel substitui os dados anteriores.

### **P: E se eu quiser guardar histórico?**
**R:** Você precisará adicionar um banco de dados (MongoDB Atlas gratuito, por exemplo). Posso ajudar se precisar!

### **P: Preciso saber programação?**
**R:** Não! Basta seguir o passo a passo acima. Copiar e colar.

### **P: A macro funciona em Mac?**
**R:** Sim, mas o Mac Excel tem limitações com VBA. Pode precisar de ajustes.

### **P: Consigo usar no celular?**
**R:** O dashboard funciona perfeitamente no celular! A macro VBA só roda no Excel desktop.

### **P: E se o servidor "adormecer"?**
**R:** Servidores gratuitos adormecem após inatividade. Ao executar a macro ou acessar a URL, ele acorda automaticamente (demora ~30 segundos na primeira vez).

---

## 🆘 Erros Comuns

### Erro: "Run-time error 70: Permission denied"

- **Solução**: Verifique se tem permissão de Internet no Excel
- No Windows: Vá em **Firewall** → Permitir o Excel

### Erro: "The remote server returned an error: (404) Not Found"

- **Solução**: URL incorreta na macro. Verifique se copiou certinho.

### Erro: "Object required"

- **Solução**: Certifique-se que as planilhas "Capa" e "Ana.XXX" existem.

### Dashboard mostra "Erro ao carregar dados"

- **Solução**: Execute a macro pelo menos uma vez para enviar dados iniciais.

---

## 📧 Precisa de Ajuda?

1. Verifique se seguiu TODOS os passos
2. Teste a URL no navegador: `https://sua-url.com/api/health`
3. Abra o console do navegador (F12) e veja os erros
4. Entre em contato com o suporte técnico

---

## ✅ Checklist Final

- [ ] Servidor deployado (Render/Railway/Vercel)
- [ ] URL copiada e salva
- [ ] Macro VBA colada no Excel
- [ ] URL da API atualizada na macro (linha 7)
- [ ] Planilha salva como `.xlsm`
- [ ] Macro testada (Alt + F8)
- [ ] Dashboard acessível no navegador
- [ ] Dados carregando corretamente

**Se marcou tudo ✅, está pronto para usar!** 🎉
