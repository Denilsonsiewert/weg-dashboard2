const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));
app.use(express.static('public'));

// Armazenamento em memória dos dados (última versão enviada)
let latestData = {
    timestamp: new Date().toISOString(),
    analises: []
};

// Endpoint para receber dados do Excel (VBA)
app.post('/api/dados', (req, res) => {
    try {
        const data = req.body;

        // Validação básica
        if (!data || !data.timestamp || !Array.isArray(data.analises)) {
            return res.status(400).json({
                success: false,
                error: 'Formato de dados inválido'
            });
        }

        // Atualiza os dados armazenados
        latestData = {
            timestamp: data.timestamp,
            analises: data.analises
        };

        console.log(`✅ Dados recebidos em ${new Date().toLocaleString('pt-BR')}`);
        console.log(`   - Total de análises: ${data.analises.length}`);

        res.json({
            success: true,
            message: 'Dados recebidos com sucesso',
            timestamp: latestData.timestamp
        });
    } catch (error) {
        console.error('❌ Erro ao processar dados:', error);
        res.status(500).json({
            success: false,
            error: 'Erro ao processar dados'
        });
    }
});

// Endpoint para o frontend buscar os dados
app.get('/api/dados', (req, res) => {
    res.json(latestData);
});

// Endpoint de health check
app.get('/api/health', (req, res) => {
    res.json({
        status: 'ok',
        timestamp: new Date().toISOString(),
        lastUpdate: latestData.timestamp
    });
});

// Rota raiz
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/public/index.html');
});

app.listen(PORT, () => {
    console.log(`🚀 Servidor rodando em http://localhost:${PORT}`);
    console.log(`📊 API disponível em /api/dados`);
    console.log(`💚 Health check em /api/health`);
});
