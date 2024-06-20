// Import dependencies
import path from 'path';
import moment from 'moment';
import exceljs from 'exceljs';
import unorm from 'unorm';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

moment.locale('pt-br');

const filePathPrimary = 'C:\\Users\\DEYVSON FELIPE\\Desktop\\busca-arquivos\\ARQUIVO.xlsx';
const filePath2024 = 'C:\\Users\\DEYVSON FELIPE\\Desktop\\busca-arquivos\\lista2024.xlsx';

const processExcelFile = async (filePath) => {
    const workbook = new exceljs.Workbook();
    let data = [];

    const getDateColumnIndex = (filePath) => {
        const fileName = path.basename(filePath);
        if (fileName === 'ARQUIVO.xlsx') {
            return 3; // Column index for ARQUIVO.xlsx
        } else if (fileName === 'lista2024.xlsx') {
            return 4; // Column index for lista2024.xlsx
        } else {
            throw new Error('Unknown file path');
        }
    };

    try {
        console.log(`Reading file: ${filePath}`);
        await workbook.xlsx.readFile(filePath);
        const worksheets = workbook.worksheets;
        const dateColumnIndex = getDateColumnIndex(filePath);
        console.log(`Using date column index: ${dateColumnIndex}`);

        worksheets.forEach(sheet => {
            sheet.eachRow((row, rowIndex) => {
                if (rowIndex === 1) return; // Skip header row
                let dateOfBirth = row.getCell(dateColumnIndex).value;
                let observation = row.getCell(5).value || ''; // Observation column index

                if (dateOfBirth instanceof Date && !isNaN(dateOfBirth.getTime())) {
                    dateOfBirth = moment.utc(dateOfBirth).format('YYYY-MM-DD');
                } else if (typeof dateOfBirth === 'string') {
                    const parsedDate = moment(dateOfBirth, [
                        'DD-MM-YYYY', 'DD/MM/YYYY', 'DD.MM.YYYY',
                        'D-M-YYYY', 'D/M/YYYY', 'D.M.YYYY',
                        'M-D-YYYY', 'M/D/YYYY', 'M.D.YYYY',
                        'YYYY-MM-DD', 'YYYY/MM/DD', 'YYYY.MM.DD'
                    ], true);

                    if (parsedDate.isValid()) {
                        dateOfBirth = parsedDate.format('YYYY-MM-DD');
                    } else {
                        dateOfBirth = 'Data Inválida';
                    }
                } else {
                    dateOfBirth = 'Data Inválida';
                }

                data.push({
                    numeracao: `R${rowIndex}`,
                    nome: row.getCell(2).value,
                    data_de_nascimento: dateOfBirth,
                    sigem: row.getCell(3).value,
                    observation: observation, // Include observation field
                    sheet: sheet.name
                });
            });
        });
    } catch (error) {
        console.error('Erro ao ler arquivo Excel:', error.message);
    }

    return data;
};

const normalizeString = (str) => {
    if (!str) return '';
    return unorm.nfd(str).replace(/[\u0300-\u036f]/g, '').toLowerCase();
};

const parseDate = (date) => {
    if (!date || date === 'Data Inválida') return 'Data Inválida';

    const parsedDate = moment.utc(date, 'YYYY-MM-DD', true);
    return parsedDate.isValid() ? parsedDate.format('DD [de] MMMM [de] YYYY') : 'Data Inválida';
};

const searchName = async (name, filePath = filePathPrimary) => {
    const normalizedQuery = normalizeString(name);
    const data = await processExcelFile(filePath);

    const results = data.filter(entry => normalizeString(entry.nome).includes(normalizedQuery))
                        .map(entry => ({
                            name: entry.nome,
                            dateOfBirth: parseDate(entry.data_de_nascimento),
                            sigem: entry.sigem,
                            observation: entry.observation,
                            sheet: entry.sheet,
                            numeracao: entry.numeracao
                        }));

    return results;
};

const searchByDate = async (date, filePath = filePathPrimary) => {
    const normalizedDate = moment.utc(date, 'YYYY-MM-DD', true).format('YYYY-MM-DD');
    const data = await processExcelFile(filePath);

    const results = data.filter(entry => entry.data_de_nascimento === normalizedDate)
                        .map(entry => ({
                            name: entry.nome,
                            dateOfBirth: parseDate(entry.data_de_nascimento),
                            sigem: entry.sigem,
                            observation: entry.observation,
                            sheet: entry.sheet,
                            numeracao: entry.numeracao
                        }));

    return results;
};

const searchSigem = async (sigem) => {
    const data = await processExcelFile(filePath2024);

    const results = data.filter(entry => entry.sigem == sigem)
                        .map(entry => ({
                            name: entry.nome,
                            dateOfBirth: parseDate(entry.data_de_nascimento),
                            sigem: entry.sigem,
                            observation: entry.observation,
                            sheet: entry.sheet,
                            numeracao: entry.numeracao
                        }));

    return results;
};

const getCellAddress = async (sheetName, numeracao, filePath) => {
    const workbook = new exceljs.Workbook();

    try {
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet(sheetName);
        const row = parseInt(numeracao.replace('R', ''), 10);
        const cell = sheet.getRow(row).getCell(1);

        return cell.address;
    } catch (error) {
        throw new Error(`Erro ao obter o endereço da célula: ${error.message}`);
    }
};

const updateEntry = async (filePath, sheetName, numeracao, newData) => {
    const workbook = new exceljs.Workbook();
    try {
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet(sheetName);
        const row = sheet.getRow(parseInt(numeracao.replace('R', ''), 10));

        if (newData.name) row.getCell(2).value = newData.name;
        if (newData.dateOfBirth) row.getCell(3).value = moment(newData.dateOfBirth, 'YYYY-MM-DD').toDate();
        if (newData.sigem) row.getCell(4).value = newData.sigem;
        if (newData.observation) row.getCell(5).value = newData.observation;

        await workbook.xlsx.writeFile(filePath);
    } catch (error) {
        throw new Error(`Erro ao atualizar a entrada: ${error.message}`);
    }
};

const deleteEntry = async (filePath, sheetName, numeracao) => {
    const workbook = new exceljs.Workbook();
    try {
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.getWorksheet(sheetName);
        const row = sheet.getRow(parseInt(numeracao.replace('R', ''), 10));
        sheet.spliceRows(row.number, 1);
        await workbook.xlsx.writeFile(filePath);
    } catch (error) {
        throw new Error(`Erro ao deletar a entrada: ${error.message}`);
    }
};

export { searchName, searchByDate, searchSigem, getCellAddress, updateEntry, deleteEntry };
