const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const app = express();
const PORT = process.env.PORT || 3000;
// Middleware
app.use(cors());
// IMPORTANTE: Aumentado o limite para 50mb devido às imagens Base64 dos gráficos
// Cada gráfico pode ter ~500KB-2MB em Base64, e com múltiplas abas pode ultrapassar 10mb
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));
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
        // Contagem de gráficos exportados
        const chartsCount = data.analises.filter(a => a.chart_image && a.chart_image.length > 0).length;
        console.log(`✅ Dados recebidos em ${new Date().toLocaleString('pt-BR')}`);
        console.log(`   - Total de análises: ${data.analises.length}`);
        console.log(`   - Gráficos exportados: ${chartsCount}`);
        res.json({
            success: true,
            message: 'Dados recebidos com sucesso',
            timestamp: latestData.timestamp,
            charts_received: chartsCount
        });
    } catch (error) {
        console.error('❌ Erro ao processar dados:', error);
        res.status(500).json({
            success: false,
            error: 'Erro ao processar dados',
            details: error.message
        });
    }
});
// Endpoint para o frontend buscar os dados
app.get('/api/dados', (req, res) => {
    res.json(latestData);
});
// Endpoint de health check
app.get('/api/health', (req, res) => {
    const chartsCount = latestData.analises.filter(a => a.chart_image && a.chart_image.length > 0).length;
    
    res.json({
        status: 'ok',
        timestamp: new Date().toISOString(),
        lastUpdate: latestData.timestamp,
        analises_count: latestData.analises.length,
        charts_count: chartsCount
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
    console.log(`📈 Suporte para imagens Base64 de gráficos (limite: 50mb)`);
});
