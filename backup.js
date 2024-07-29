import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import axios from 'axios';
import FormData from 'form-data';
import OpenAI from "openai";
import ExcelJS from 'exceljs';

// Necessário para resolver __dirname com ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Configurar o armazenamento dos arquivos com multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, path.join(__dirname, 'uploads/'));
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + '-' + Date.now() + path.extname(file.originalname));
    }
});

const upload = multer({ storage: storage });

// Middleware para servir arquivos estáticos (opcional)
app.use(express.static(path.join(__dirname, 'public')));
const openaiApiKey = 'sk-proj-qB5osZlQytxfbGNicWMpT3BlbkFJ4iB9n9Yc0mbnota4fYd7'; // Substitua pelo seu OpenAI API Key
const pdfCoApiKey = 'davi.catarino@hotmail.com_3z3IcIhnniVkf7hfP9wv280wAnh0q4Uv15cEGf3Gufte94qea6nHZw7v4vXDs3w7'; // Substitua pelo seu PDF.co API Key

const openai = new OpenAI({
    apiKey: openaiApiKey
});

function isValidJSON(jsonString) {
    try {
        JSON.parse(jsonString);
        return true;
    } catch (e) {
        return false;
    }
}

async function transformTextToStructuredData(text) {
    const response = await openai.chat.completions.create({
        messages: [{
            role: "system",
            content: `Você é um assistente que transforma textos de extratos bancários em dados estruturados. Sua tarefa é extrair todos os lançamentos de compra e saque, lançamentos de produtos e serviços, outros lançamentos, e as compras parceladas. Para cada lançamento, extraia a data, nome do estabelecimento, valor e a quantidade de parcelas.

            Instruções detalhadas:
            1. Para lançamentos que não possuem o número de parcelas ao lado do nome do estabelecimento, considere a quantidade de parcelas como "0/0".
            2. Para lançamentos que possuem o número de parcelas ao lado do nome do estabelecimento, informe corretamente as parcelas.
            3. Verifique duas vezes se as parcelas estão sendo extraídas corretamente.
            4. Retorne os dados em um único objeto JSON com quatro arrays: "data", "estabelecimento", "valor" e "N_de_parcela".

            Aqui está o texto do extrato bancário para ser analisado:
            \n\nTexto:\n${text}`

        }],
        model: "gpt-4o",
        temperature: 0.5,
        response_format: { "type": "json_object" }
    });

    console.log('chegou aq:aaaaaaaaaaaaaaaaaaa');
    let structuredData = response.choices[0].message.content.trim();

    // Adiciona depuração
    console.log('Structured Data:', structuredData);

    if (!isValidJSON(structuredData)) {
        structuredData += '}';
        if (!isValidJSON(structuredData)) {
            console.error('Erro ao parsear o JSON: JSON ainda está incompleto');
            return null;
        }
    }

    try {
        return JSON.parse(structuredData);
    } catch (error) {
        console.error('Erro ao parsear o JSON:', error);
        return null;
    }
}

async function addDataToExcel(data, filename) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(1);

    // Adiciona uma linha de cabeçalho se não existir
    if (worksheet.actualRowCount === 0) {
        worksheet.addRow(['Data', 'estabelecimento', 'Valor', 'N_de_Parcelas']);
    }

    // Adiciona as novas linhas diretamente
    for (let i = 0; i < data.data.length; i++) {
        worksheet.addRow([data.data[i], data.estabelecimento[i], data.valor[i], data.N_de_parcela[i]]);
    }
    console.log(data); // Adiciona depuração

    await workbook.xlsx.writeFile(filename);
    console.log('Data added to Excel file successfully'); // Adiciona depuração
}

// Rota para upload de arquivo
app.post('/upload', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).send('Nenhum arquivo foi enviado.');

    const filePath = req.file.path;
    const excelFilename = path.join(__dirname, './planilha/teste.xlsx'); // Substitua pelo nome do seu arquivo Excel

    try {
        const fileContent = fs.readFileSync(filePath);
        const formData = new FormData();
        formData.append('file', fileContent, {
            filename: req.file.originalname,
            contentType: 'application/pdf'
        });

        const uploadResponse = await axios.post('https://api.pdf.co/v1/file/upload', formData, {
            headers: { ...formData.getHeaders(), 'x-api-key': pdfCoApiKey }
        });

        const fileUrl = uploadResponse.data.url;
        const conversionData = { url: fileUrl, pages: "", lang: "eng" };
        const conversionResponse = await axios.post('https://api.pdf.co/v1/pdf/convert/to/text', conversionData, {
            headers: { 'Content-Type': 'application/json', 'x-api-key': pdfCoApiKey }
        });

        const textFileUrl = conversionResponse.data.url; // Obtém a URL do arquivo de texto

        // Faz uma requisição HTTP para obter o conteúdo do arquivo de texto
        const textResponse = await axios.get(textFileUrl);
        const extractedText = textResponse.data;
        console.log('Extracted Text:', extractedText); // Adiciona depuração
        const structuredData = await transformTextToStructuredData(extractedText);

        if (structuredData) {
            await addDataToExcel(structuredData, excelFilename);
            res.send('Dados adicionados à planilha com sucesso.');
        } else {
            res.status(500).send('Erro ao processar os dados estruturados.');
        }
    } catch (error) {
        console.error('Erro ao processar o arquivo:', error);
        res.status(500).send('Erro ao processar o arquivo.');
    } finally {
        fs.unlinkSync(filePath); // Remover o arquivo local após o upload
    }
});

// Iniciar o servidor
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
