import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { searchName, searchByDate, searchSigem, getCellAddress } from './dataProcessor.js';
import exceljs from 'exceljs'; // Importe o pacote exceljs
import moment from 'moment'; // Importe o pacote moment para manipulação de datas
import { exec } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

// Middleware para servir arquivos estáticos
app.use(express.static(path.join(__dirname, '../public')));

// Middleware para parsear JSON
app.use(express.json());

// Endpoint para pesquisar no arquivo principal
app.post('/search', async (req, res) => {
    const { name, dateOfBirth } = req.body;
    let results = [];

    if (name) {
        results = await searchName(name);
    } else if (dateOfBirth) {
        results = await searchByDate(dateOfBirth);
    }

    res.json(results);
});

// Endpoint para abrir o Excel principal
app.get('/open-excel', async (req, res) => {
    const { sheet, numeracao } = req.query;
    try {
        const cellAddress = await getCellAddress(sheet, numeracao, 'ARQUIVO.xlsx');
        const filePath = path.join(__dirname, '../ARQUIVO.xlsm');
        const scriptPath = path.join(__dirname, 'open_excel.ps1');

        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -filePath "${filePath}" -sheet "${sheet}" -cell "${cellAddress}"`;

        exec(command, (error, stdout, stderr) => {
            if (error) {
                return res.status(500).send('Error opening Excel');
            }
            console.log(`stdout: ${stdout}`);
            console.error(`stderr: ${stderr}`);
            res.send('Excel opened');
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error opening Excel');
    }
});

// Endpoint para editar um aluno na lista de 2024
app.put('/edit-student', async (req, res) => {
    const { name, dateOfBirth, sigem, observation, numeracao, sheet } = req.body;
    try {
        await editStudent(name, dateOfBirth, sigem, observation, numeracao, sheet, 'lista2024.xlsx');
        res.send('Student updated');
    } catch (error) {
        console.error(error);
        res.status(500).send('Error editing student');
    }
});

// Endpoint DELETE para remover o aluno de ambos os arquivos
app.delete('/delete-student', async (req, res) => {
    const { numeracao, sheet } = req.query;
    
    try {
        await deleteStudent(numeracao, sheet, 'lista2024.xlsx');
        await deleteStudent(numeracao, sheet, 'lista2024.xlsm'); // Exclui de ambos os arquivos
        res.send('Aluno excluído de todos os arquivos');
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao excluir aluno');
    }
});

// Função para editar um aluno
async function editStudent(name, dateOfBirth, sigem, observation, numeracao, sheetName, filePath) {
    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    const row = parseInt(numeracao.replace('R', ''), 10);

    const dateColumnIndex = 4; // Ajustar conforme sua lógica de negócio
    sheet.getRow(row).getCell(2).value = name;
    sheet.getRow(row).getCell(dateColumnIndex).value = moment(dateOfBirth, 'YYYY-MM-DD').toDate();
    sheet.getRow(row).getCell(3).value = sigem;
    sheet.getRow(row).getCell(5).value = observation; // Ajustar conforme sua lógica de negócio

    await workbook.xlsx.writeFile(filePath);
}

// Função para excluir aluno
async function deleteStudent(numeracao, sheetName, filePath) {
    if (!fs.existsSync(filePath)) {
        throw new Error(`Arquivo não encontrado: ${filePath}`);
    }

    const workbook = new exceljs.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(sheetName);
    const row = parseInt(numeracao.replace('R', ''), 10);
    sheet.spliceRows(row, 1); 
    await workbook.xlsx.writeFile(filePath);
}

// Endpoint para buscar as notas do aluno
app.get('/student-grades', async (req, res) => {
    const { name, dateOfBirth } = req.query;
    try {
        const grades = await getGrades(name, dateOfBirth);
        res.json(grades);
    } catch (error) {
        console.error(error);
        res.status(500).send('Erro ao buscar notas');
    }
});

// Função para buscar as notas
async function getGrades(name, dateOfBirth) {
    // Lógica para buscar as notas do aluno baseado no nome e data de nascimento
    return [
        { subject: 'Matemática', grade: 85 },
        { subject: 'Português', grade: 90 }
    ];
}

// Endpoint para pesquisar na lista de 2024
app.post('/search2024', async (req, res) => {
    const { name, dateOfBirth, sigem } = req.body;
    let results = [];

    if (name) {
        results = await searchName(name, 'lista2024.xlsx');
    } else if (dateOfBirth) {
        results = await searchByDate(dateOfBirth, 'lista2024.xlsx');
    } else if (sigem) {
        results = await searchSigem(sigem);
    }

    res.json(results);
});

// Endpoint para abrir o Excel da lista de 2024
app.get('/open-excel-2024', async (req, res) => {
    const { sheet, numeracao } = req.query;
    try {
        const cellAddress = await getCellAddress(sheet, numeracao, 'lista2024.xlsx');
        const filePath = path.join(__dirname, '../lista2024.xlsm');
        const scriptPath = path.join(__dirname, 'open_excel.ps1');

        const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -filePath "${filePath}" -sheet "${sheet}" -cell "${cellAddress}"`;

        exec(command, (error, stdout, stderr) => {
            if (error) {
                console.error(`exec error: ${error}`);
                return res.status(500).send('Error opening Excel');
            }
            console.log(`stdout: ${stdout}`);
            console.error(`stderr: ${stderr}`);
            res.send('Excel opened');
        });
    } catch (error) {
        console.error(error);
        res.status(500).send('Error opening Excel');
    }
});

// Inicia o servidor Node.js
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
