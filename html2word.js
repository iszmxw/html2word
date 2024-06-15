const fs = require('fs');
const path = require('path');
const htmlDocx = require('html-docx-js');
const { Buffer } = require('buffer');
const util = require('util');

// 使用 promisify 将 fs.readFile 和 fs.writeFile 转换为返回 Promise 的函数
const readFile = util.promisify(fs.readFile);
const writeFile = util.promisify(fs.writeFile);

const HTML_FILE_PATH = './1.html'; // 目标富文本内容文件
const TEMPLATE_FILE_PATH = './template.docx'; // 空模板文件
const OUTPUT_FILE_PATH = '2.docx'; // 输出结果文件

async function main() {
    try {
        // 读取 HTML 文件内容
        const htmlContent = await readHtmlFile(HTML_FILE_PATH);
        // 渲染并保存 Word 文档
        await renderWordDocument(htmlContent, TEMPLATE_FILE_PATH, OUTPUT_FILE_PATH);
    } catch (error) {
        console.error('处理文件时出错: ', error.message);
    }
}

// 读取 HTML 文件内容的函数
async function readHtmlFile(filePath) {
    try {
        // 异步读取文件内容
        const data = await readFile(filePath, 'utf8');
        return data;
    } catch (error) {
        // 抛出错误以便在调用方捕获
        throw new Error(`读取文件 ${filePath} 时出错: ${error.message}`);
    }
}

// 渲染并保存 Word 文档的函数
async function renderWordDocument(htmlContent, templatePath, outputPath) {
    try {
        // 将 HTML 内容转换为二进制 Blob
        const htmlBuffer = htmlDocx.asBlob(htmlContent);
        // 将 Blob 转换为 ArrayBuffer
        const arrayBuffer = await htmlBuffer.arrayBuffer();
        // 将 ArrayBuffer 转换为 Node.js 的 Buffer
        const buffer = Buffer.from(arrayBuffer);
        // 异步写入文件
        await writeFile(outputPath, buffer);
        console.log('文件已成功保存!');
    } catch (error) {
        throw new Error(`创建 Word 文档时出错: ${error.message}`);
    }
}

// 执行主函数
main();
