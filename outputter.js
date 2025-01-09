// 1111.mjs 或 1111.js（取决于是否更改文件扩展名）
import axios from 'axios';
import ExcelJS from 'exceljs';
import pLimit from 'p-limit';
import { promises as fs } from 'fs';
import path from 'path';

// 配置参数
const TOTAL_REQUESTS = 5500;          // 总请求次数
const CONCURRENT_LIMIT = 5;            // 并发限制，根据需要调整
const MAX_RETRIES = 5;                 // 最大重试次数
const INITIAL_DELAY = 1000;            // 初始延迟（毫秒）
const SAVE_INTERVAL = 30;             // 每100次请求保存一次
const EXCEL_FILE_PATH = path.resolve('./user_data.xlsx'); // Excel文件路径

// 创建Excel工作簿和工作表
const workbook = new ExcelJS.Workbook();
let worksheet;

// 数据缓存
let dataBuffer = [];
let completedRequests = 0;

// 初始化工作簿和工作表
async function initializeWorkbook() {
    try {
        // 检查Excel文件是否存在
        await fs.access(EXCEL_FILE_PATH);
        console.log('Excel file exists. Loading...');
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        worksheet = workbook.getWorksheet('User Data');
        if (!worksheet) {
            // 如果工作表不存在，则创建并添加表头
            worksheet = workbook.addWorksheet('User Data');
            addHeader();
        }
    } catch (err) {
        if (err.code === 'ENOENT') {
            console.log('Excel file does not exist. Creating a new one...');
            // 创建新的工作表并添加表头
            worksheet = workbook.addWorksheet('User Data');
            addHeader();
        } else {
            console.error('Error accessing Excel file:', err);
            process.exit(1);
        }
    }
}

// 添加表头到工作表
function addHeader() {
    const header = [
        'Country', 
        'First Name', 
        'Last Name', 
        'City', 
        'Address', 
        'Zip Code', 
        'Date of Birth', 
        'Phone', 
        'Bank IBAN', 
        'Credit Card Number', 
        'Credit Card Expiration Date', 
        'Credit Card CVV2'
    ];
    worksheet.addRow(header);
}

// 发送单次请求的函数，包含重试逻辑
async function fetchSingleData(index, retries = 0) {
    try {
        const response = await axios.get('https://outputter.io/api/identity/DE', {
            headers: {
                'Authorization': 'Basic ' + Buffer.from('8159:d040295be097c0d62814df96a808edbf').toString('base64')
            },
            timeout: 10000 // 设置超时时间为10秒
        });
        console.log(`Request ${index + 1} completed`);
        return response.data;  // 返回API响应的数据
    } catch (error) {
        if (error.response && error.response.status === 429) {
            // 获取 Retry-After 头（如果有）
            const retryAfter = error.response.headers['retry-after'];
            const delayTime = retryAfter ? parseInt(retryAfter) * 1000 : INITIAL_DELAY * Math.pow(2, retries);
            if (retries < MAX_RETRIES) {
                console.warn(`Request ${index + 1} received 429. Retrying after ${delayTime} ms... (Retry ${retries + 1}/${MAX_RETRIES})`);
                await delay(delayTime);
                return fetchSingleData(index, retries + 1);
            } else {
                console.error(`Request ${index + 1} failed after ${MAX_RETRIES} retries.`);
                return { status: 'ERROR', message: 'Max retries reached' };
            }
        } else {
            console.error(`Error during request ${index + 1}:`, error.message);
            return { status: 'ERROR', message: error.message };
        }
    }
}

// 延迟函数，用于控制重试等待时间
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// 写入数据到Excel
async function writeToExcel() {
    if (dataBuffer.length === 0) return;
    console.log(`Saving ${dataBuffer.length} records to Excel...`);
    dataBuffer.forEach(item => {
        worksheet.addRow([
            item.country,
            item.firstName,
            item.lastName,
            item.city,
            item.address,
            item.zipCode,
            item.dateOfBirth,
            item.phone,
            item.bankIban,
            item.creditCardNumber,
            item.creditCardExpirationDate,
            item.creditCardCVV2
        ]);
    });
    dataBuffer = []; // 清空缓存
    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
    console.log(`Saved to ${EXCEL_FILE_PATH}`);
}

// 处理程序退出时保存数据
async function handleExit() {
    console.log('\nProcess exiting. Saving remaining data...');
    await writeToExcel();
    process.exit();
}

// 主函数，发起请求
async function fetchData() {
    await initializeWorkbook();

    const limit = pLimit(CONCURRENT_LIMIT);
    const tasks = [];

    for (let i = 0; i < TOTAL_REQUESTS; i++) {
        tasks.push(
            limit(() => fetchSingleData(i).then(data => {
                completedRequests++;
                if (data.status === 'OK') {
                    dataBuffer.push({
                        country: data.response.country,
                        firstName: data.response.firstName,
                        lastName: data.response.lastName,
                        city: data.response.city,
                        address: data.response.address,
                        zipCode: data.response.zipCode,
                        dateOfBirth: data.response.dateOfBirth,
                        phone: data.response.phone,
                        bankIban: data.response.bank.iban,
                        creditCardNumber: data.response.creditCard.number,
                        creditCardExpirationDate: data.response.creditCard.expirationDate,
                        creditCardCVV2: data.response.creditCard.cvv2
                    });
                } else {
                    console.warn(`Skipping request ${i + 1} due to error: ${data.message}`);
                }

                // 定期保存数据
                if (completedRequests % SAVE_INTERVAL === 0) {
                    writeToExcel();
                }
            }))
        );
    }

    try {
        await Promise.all(tasks);
        console.log('All requests completed.');
        await writeToExcel(); // 保存剩余的数据
    } catch (error) {
        console.error('Error processing requests:', error);
    }
}

// 监听进程终止信号，确保数据保存
process.on('SIGINT', handleExit);
process.on('SIGTERM', handleExit);

// 执行主函数
fetchData();
