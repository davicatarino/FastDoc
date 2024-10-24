import express from 'express';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import OpenAI from "openai";
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import pdfParse from 'pdf-parse';
import { PDFDocument } from 'pdf-lib'; // Importando a biblioteca pdf-lib

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

const openaiApiKey = process.env.GPT_KEY; // Substitua pelo seu OpenAI API Key
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
async function transformTextToFormatText(text) {
    const response = await openai.chat.completions.create({
        messages: [{
            role: "system",
            content: `Você é um assistente especializado em limpar textos extraídos de PDFs de extratos bancários.
                  1. Formate e imprima somente os lançamentos da fatura do cartão.
                  2. Remova todas as informações que não sejam lançamentos da fatura do cartão.
                  3. exemplo de dado a ser extraido : "07/06KABUM             12/12408,00". "07/06" = data, "kabum" = estabelecimento, "12/12" = numero de parcelas(12 parcelas de 12), "408,00" = valor. 
                  4. Ignore os lançamentos que aparecem após "Compras parceladas - próximas faturas".

            \n\nTexto:\n${text}`
        }],
        model: "gpt-4o",
        temperature: 0.2,

    });

    let formatData = response.choices[0].message.content.trim();
    console.log(formatData)

    try {
        return formatData;
    } catch (error) {
        console.error('Erro ao parsear o JSON:', error);
        return null;
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
            5. Ignore os lançamentos após o "Compras parceladas - próximas faturas". 

            exemplo de dado a ser extraido : "07/06KABUM             12/12408,00". "07/06" = data, "kabum" = estabelecimento, "12/12" = numero de parcelas(12 parcelas de 12), "408,00" = valor. 

            Aqui está o texto do extrato bancário para ser analisado:
            \n\nTexto:\n${text}`
        }],
        model: "gpt-4o",
        max_tokens: 4096,
        temperature: 0.2,
        response_format: { "type": "json_object" }
    });

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

function preprocessTextWithLineSpaces(text) {
    // Adiciona uma linha de espaço entre cada linha
    text = text.replace(/([^\n])\n([^\n])/g, '$1\n\n$2');
    return text;
}

async function extractTextFromPdf(pdfPath) {
    const dataBuffer = await fs.readFile(pdfPath);
    const pdfDoc = await PDFDocument.load(dataBuffer);

    const totalPages = pdfDoc.getPageCount();
    if (totalPages <= 4) {
        throw new Error('PDF não possui páginas suficientes para remover a primeira e as duas últimas páginas.');
    }

    const pagesToKeep = Array.from({ length: totalPages - 4 }, (_, i) => i + 1);
    const newPdfDoc = await PDFDocument.create();

    for (const pageIndex of pagesToKeep) {
        const [page] = await newPdfDoc.copyPages(pdfDoc, [pageIndex]);
        newPdfDoc.addPage(page);
    }

    const newPdfBytes = await newPdfDoc.save();
    const newPdfBuffer = Buffer.from(newPdfBytes);

    const newPdfData = await pdfParse(newPdfBuffer);
    const extractedText = newPdfData.text;

    // Pré-processa o texto extraído
    const processedText = preprocessTextWithLineSpaces(extractedText);
    return processedText;
}

async function addDataToExcel(data, filename) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(1);

    // Adiciona uma linha de cabeçalho se não existir
    if (worksheet.actualRowCount === 0) {
        worksheet.addRow(['Data', 'Estabelecimento', 'Valor', 'N_de_Parcelas']);
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
        const extractedText = await extractTextFromPdf(filePath);
        console.log('Extracted Text:', extractedText); // Adiciona depuração
        const formatData = await transformTextToFormatText(extractedText);
        const structuredData = await transformTextToStructuredData(formatData);

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
        await fs.unlink(filePath); // Remover o arquivo local após o upload
    }
});

// Iniciar o servidor
app.listen(PORT, () => console.log(`Servidor rodando na porta ${PORT}`));
