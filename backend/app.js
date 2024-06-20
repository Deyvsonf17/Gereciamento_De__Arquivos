import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import exceljs from 'exceljs';
import { exec } from 'child_process';
import { searchName, searchByDate, searchSigem, getCellAddress, deleteEntry } from './dataProcessor.js';

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
                console.error(`exec error: ${error}`);
                return res.status(500).send('Error opening Excel');
            }
            console.log(`stdout: ${stdout}`);
            console.error(`stderr: ${stderr}`);
            res.send('Excel opened');
        });
    } catch (error) {
        console.error('Error in /open-excel endpoint:', error.message);
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
        console.error('Error in /edit-student endpoint:', error.message);
        res.status(500).send('Error editing student');
    }
});

// Endpoint para excluir um aluno da lista de 2024
app.delete('/delete-student', async (req, res) => {
    const { numeracao, sheet } = req.query;
    const xlsmFilePath = path.join(__dirname, '../lista2024.xlsm');
    const xlsxFilePath = path.join(__dirname, '../lista2024.xlsx');
    const scriptPath = path.join(__dirname, 'open_excel.ps1');

    try {
        // Exclui do arquivo .xlsx
        await deleteEntry(xlsxFilePath, sheet, numeracao);

        // Copia o arquivo .xlsx para .xlsm e exclui o .xlsm anterior
        exec(`copy "${xlsxFilePath}" "${xlsmFilePath}"`, async (error, stdout, stderr) => {
            if (error) {
                console.error(`exec error: ${error}`);
                return res.status(500).send('Error copying Excel files');
            }
            console.log(`stdout: ${stdout}`);
             console.error(`stderr: ${stderr}`);

            const cellAddress = await getCellAddress(sheet, numeracao, 'lista2024.xlsx');
            const command = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -filePath "${xlsmFilePath}" -sheet "${sheet}" -cell "${cellAddress}" -delete`;

            exec(command, (error, stdout, stderr) => {
                if (error) {
                    console.error(`exec error: ${error}`);
                    return res.status(500).send('Error deleting student in Excel');
                }
                console.log(`stdout: ${stdout}`);
                console.error(`stderr: ${stderr}`);
                res.send('Student deleted');
            });
        });
    } catch (error) {
        console.error('Error in /delete-student endpoint:', error.message);
        res.status(500).send('Error deleting student');
    }
});

// Endpoint para buscar as notas do aluno
app.get('/student-grades', async (req, res) => {
    const { name, dateOfBirth } = req.query;
    try {
        const grades = await getGrades(name, dateOfBirth);
        res.json(grades);
    } catch (error) {
        console.error('Error in /student-grades endpoint:', error.message);
        res.status(500).send('Erro ao buscar notas');
    }
});

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
        console.error('Error in /open-excel-2024 endpoint:', error.message);
        res.status(500).send('Error opening Excel');
    }
});

// Inicia o servidor Node.js
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
