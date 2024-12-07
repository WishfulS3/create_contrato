import { bitable } from '@lark-base-open/js-sdk';
import './index.scss';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';

async function main() {
    const recordContentElement = document.getElementById('recordContent');
    const uploadTemplateButton = document.getElementById('uploadTemplate');
    const generateContractButton = document.getElementById('generateContract');
    const templateFileInput = document.getElementById('templateFile') as HTMLInputElement;

    let selectedRecord: any = null;
    let templateContent: ArrayBuffer | null = null;

    async function updateSelectedRecord() {
        try {
            const selection = await bitable.base.getSelection();
            if (selection.tableId && selection.recordId) {
                const table = await bitable.base.getTableById(selection.tableId);
                selectedRecord = await table.getRecordById(selection.recordId);
                if (recordContentElement) {
                    recordContentElement.innerHTML = `<pre>${JSON.stringify(selectedRecord, null, 2)}</pre>`;
                }
            } else {
                selectedRecord = null;
                if (recordContentElement) {
                    recordContentElement.textContent = '请在表格中选择一条记录';
                }
            }
        } catch (error) {
            console.error('Error updating selected record:', error);
            if (recordContentElement) {
                recordContentElement.textContent = '获取记录失败，请重试';
            }
        }
    }

    await updateSelectedRecord();
    bitable.base.onSelectionChange(updateSelectedRecord);

    uploadTemplateButton?.addEventListener('click', async () => {
        const file = templateFileInput.files?.[0];
        if (!file) {
            alert('请先选择模板文件');
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            templateContent = e.target?.result as ArrayBuffer;
            console.log('模板上传成功');
            alert('模板上传成功');
        };
        reader.onerror = () => {
            console.error('模板上传失败');
            alert('模板上传失败');
        };
        reader.readAsArrayBuffer(file);
    });

    generateContractButton?.addEventListener('click', async () => {
        if (!selectedRecord) {
            alert('请先选择一条记录');
            return;
        }
        if (!templateContent) {
            alert('请先上传模板文件');
            return;
        }

        try {
            const table = await bitable.base.getActiveTable();
            const fields = await table.getFieldMetaList();
            const fieldMap = new Map(fields.map(field => [field.id, field.name]));

            console.log('Available fields:', Object.fromEntries(fieldMap));

            const zip = new PizZip(templateContent);
            const doc = new Docxtemplater(zip, { 
                paragraphLoop: true, 
                linebreaks: true,
                delimiters: {
                    start: '{{',
                    end: '}}'
                }
            });

            const flattenedRecord = flattenRecord(selectedRecord.fields, fieldMap);
            console.log('Flattened record:', flattenedRecord);

            doc.setData(flattenedRecord);
            doc.render();

            const out = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });

            // 创建文件对象
            const file = new File([out], 'generated_contract.docx', {
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });

            // 上传文件
            const [fileToken] = await bitable.base.batchUploadFile([file]);
            console.log('Uploaded file token:', fileToken);

            // 找到名为 'file' 的字段
            const fileField = fields.find(field => field.name === 'file');
            if (!fileField) {
                throw new Error('未找到名为 "file" 的字段');
            }

            // 获取当前选中的记录ID
            const selection = await bitable.base.getSelection();
            if (!selection.recordId) {
                throw new Error('未选中任何记录');
            }

            // 更新记录
            await table.setCellValue(fileField.id, selection.recordId, [{
                name: file.name,
                size: file.size,
                type: file.type,
                token: fileToken,
                timeStamp: Date.now(),
            }]);

            console.log('合同生成成功并保存到记录中');
            alert('合同生成成功并保存到记录中');
        } catch (error) {
            console.error('Error generating contract:', error);
            alert('生成合同失败，请检查模板格式是否正确');
        }
    });
}

function flattenRecord(record: any, fieldMap: Map<string, string>): Record<string, string> {
    const flattened: Record<string, string> = {};

    for (const [key, value] of Object.entries(record)) {
        let fieldName = fieldMap.get(key) || key;
        // 保留中文字符，只替换其他特殊字符
        fieldName = fieldName.replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, '_');

        if (fieldName === 'file') continue; // 跳过 file 字段

        if (Array.isArray(value) && value.length > 0) {
            if (typeof value[0] === 'object' && 'text' in value[0]) {
                flattened[fieldName] = value[0].text;
            } else if (typeof value[0] === 'object' && 'name' in value[0]) {
                flattened[fieldName] = value[0].name;
            } else {
                flattened[fieldName] = JSON.stringify(value);
            }
        } else if (value === null) {
            flattened[fieldName] = '';
        } else {
            flattened[fieldName] = String(value);
        }
    }

    return flattened;
}

main().catch(console.error);