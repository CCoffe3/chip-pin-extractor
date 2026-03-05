// 全局变量，存储提取的引脚信息
let pinData = [];

// 页面加载完成后初始化事件监听
document.addEventListener('DOMContentLoaded', function() {
    // 绑定按钮事件
    document.getElementById('extractBtn').addEventListener('click', extractPinInfo);
    document.getElementById('sortBtn').addEventListener('click', smartSortPins); // 使用智能排序
    document.getElementById('exportBtn').addEventListener('click', exportToExcel);
    
    // 初始化拖拽功能
    initDragAndDrop();
    
    // 绑定自动排列和降序排列复选框的 change 事件
    const autoArrangeCheckbox = document.getElementById('autoArrange');
    const reverseCheckbox = document.getElementById('reverseBottomTop');
    
    if (autoArrangeCheckbox) {
        autoArrangeCheckbox.addEventListener('change', function() {
            // 重新应用 Position 分配
            if (pinData.length > 0) {
                reapplyPositionAllocation();
            }
        });
    }
    
    if (reverseCheckbox) {
        reverseCheckbox.addEventListener('change', function() {
            // 重新应用 Position 分配
            if (pinData.length > 0) {
                reapplyPositionAllocation();
            }
        });
    }
});

// 初始化拖拽功能
function initDragAndDrop() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    
    // 拖拽事件
    dropZone.addEventListener('dragover', function(e) {
        e.preventDefault();
        dropZone.classList.add('active');
    });
    
    dropZone.addEventListener('dragleave', function() {
        dropZone.classList.remove('active');
    });
    
    dropZone.addEventListener('drop', function(e) {
        e.preventDefault();
        dropZone.classList.remove('active');
        
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
        }
    });
    
    // 点击事件
    dropZone.addEventListener('click', function() {
        fileInput.click();
    });
}

// 全局变量，存储芯片信息
let chipInfo = {
    totalPins: 0
};

// 从图片中提取文本
function extractTextFromImage(image) {
    return new Promise((resolve, reject) => {
        Tesseract.recognize(
            image,
            'eng+chi_sim', // 支持英文和简体中文
            {
                logger: info => {
                    console.log('OCR 进度:', info);
                    // 更新进度显示
                    const progressDiv = document.getElementById('ocrProgress');
                    if (progressDiv && info.status === 'recognizing text') {
                        progressDiv.textContent = `OCR 识别进度：${Math.round(info.progress * 100)}%`;
                    }
                },
                tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_-/+.:*', // 允许的字符集
            }
        ).then(result => {
            console.log('OCR 识别结果:', result.data.text);
            resolve(result.data.text);
        }).catch(error => {
            console.error('OCR 识别失败:', error);
            reject(error);
        });
    });
}

// 解析PDF页码范围
function parsePageRange(rangeString) {
    if (!rangeString) return [];
    
    const pages = new Set();
    const ranges = rangeString.split(',');
    
    for (const range of ranges) {
        const parts = range.trim().split('-');
        if (parts.length === 1) {
            // 单个页码
            const page = parseInt(parts[0]);
            if (!isNaN(page)) {
                pages.add(page);
            }
        } else if (parts.length === 2) {
            // 页码范围
            const start = parseInt(parts[0]);
            const end = parseInt(parts[1]);
            if (!isNaN(start) && !isNaN(end)) {
                for (let i = start; i <= end; i++) {
                    pages.add(i);
                }
            }
        }
    }
    
    return Array.from(pages).sort((a, b) => a - b);
}

// 从 PDF 中提取文本
function extractTextFromPDF(pdfFile, pageRange) {
    return new Promise((resolve, reject) => {
        const fileReader = new FileReader();
        fileReader.onload = function() {
            const typedArray = new Uint8Array(this.result);
            pdfjsLib.getDocument(typedArray).promise.then(pdf => {
                let text = '';
                const pagePromises = [];
                
                // 确定要处理的页码
                const pagesToProcess = pageRange.length > 0 ? pageRange : Array.from({length: pdf.numPages}, (_, i) => i + 1);
                const totalPages = pagesToProcess.length;
                let processedPages = 0;
                
                // 更新进度
                const updateProgress = () => {
                    const progressDiv = document.getElementById('pdfProgress');
                    if (progressDiv) {
                        progressDiv.textContent = `PDF 处理进度：${Math.round((processedPages / totalPages) * 100)}% (${processedPages}/${totalPages})`;
                    }
                };
                
                for (const pageNum of pagesToProcess) {
                    if (pageNum >= 1 && pageNum <= pdf.numPages) {
                        pagePromises.push(pdf.getPage(pageNum).then(page => {
                            return page.getTextContent().then(content => {
                                // 按 y 坐标排序，然后按 x 坐标排序，以获得更准确的文本顺序
                                const items = content.items.map(item => ({
                                    str: item.str,
                                    x: item.transform[4],
                                    y: item.transform[5]
                                }));
                                
                                // 先按 y 坐标分组（同一行的文本）
                                const tolerance = 5; // y 坐标容差
                                const lines = [];
                                let currentLine = [];
                                let lastY = null;
                                
                                items.sort((a, b) => b.y - a.y); // 从上到下排序
                                
                                items.forEach(item => {
                                    if (lastY === null || Math.abs(item.y - lastY) > tolerance) {
                                        if (currentLine.length > 0) {
                                            // 按 x 坐标排序当前行
                                            currentLine.sort((a, b) => a.x - b.x);
                                            lines.push(currentLine.map(i => i.str).join(' '));
                                            currentLine = [];
                                        }
                                        lastY = item.y;
                                    }
                                    currentLine.push(item);
                                });
                                
                                if (currentLine.length > 0) {
                                    currentLine.sort((a, b) => a.x - b.x);
                                    lines.push(currentLine.map(i => i.str).join(' '));
                                }
                                
                                text += lines.join('\n') + ' ';
                                processedPages++;
                                updateProgress();
                            });
                        }));
                    }
                }
                
                Promise.all(pagePromises).then(() => {
                    resolve(text);
                }).catch(error => {
                    reject(error);
                });
            }).catch(error => {
                reject(error);
            });
        };
        fileReader.readAsArrayBuffer(pdfFile);
    });
}

// 从文本中提取引脚信息
function extractPinInfoFromText(text) {
    const pins = [];
    const seenNumbers = new Set();
    
    console.log('开始从文本中提取引脚信息...');
    console.log('文本长度:', text.length);
    
    // 按行处理文本
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    console.log('总行数:', lines.length);
    
    // 尝试 1：匹配表格格式（Pin # 在中间列）
    // 格式：Pin Name | I/O | Pin # | Description
    // 例如：NC - 1 NC
    // 例如：VCC_OK_ON DP 6 Core Power
    for (let i = 0; i < lines.length; i++) {
        const trimmedLine = lines[i].trim();
        
        // 跳过标题行和空行
        if (trimmedLine.includes('Pin#') || trimmedLine.includes('Pin Name') || 
            trimmedLine.includes('Description') || trimmedLine.length < 3) {
            continue;
        }
        
        // 分割行为多个部分
        const parts = trimmedLine.split(/\s+/).filter(p => p.length > 0);
        
        if (parts.length >= 3) {
            // 寻找数字（Pin #）
            let pinNumber = -1;
            let pinNameParts = [];
            
            for (let j = 0; j < parts.length; j++) {
                const part = parts[j];
                const numMatch = part.match(/^(\d+)$/);
                
                if (numMatch && pinNumber === -1) {
                    // 找到第一个纯数字作为 Pin #
                    pinNumber = parseInt(part);
                    break;
                } else {
                    // 在找到 Pin # 之前的部分都是 Pin Name
                    pinNameParts.push(part);
                }
            }
            
            // 如果找到了 Pin #，提取 Pin Name
            if (pinNumber > 0 && pinNumber <= 1000 && pinNameParts.length > 0) {
                const pinName = pinNameParts.join(' ');
                
                // 验证引脚名称
                if (isValidPinName(pinName)) {
                    console.log(`行 ${i + 1}: Pin #="${pinNumber}", Pin Name="${pinName}"`);
                    
                    if (!seenNumbers.has(pinNumber)) {
                        seenNumbers.add(pinNumber);
                        
                        pins.push({
                            Number: pinNumber,
                            Name: pinName,
                            Type: 'Passive', // 默认值设为 Passive
                            Visible: 'TRUE',
                            Shape: 'Line',
                            PinGroup: '', // 保持空白
                            Position: '', // 由自动排列功能分配
                            Section: 'A'
                        });
                    }
                }
            }
        }
    }
    
    // 如果没有从表格格式中提取到数据，尝试方法 2：Pin # 在第一列
    if (pins.length === 0) {
        console.log('表格格式未识别到数据，尝试方法 2（Pin # 在第一列）...');
        
        for (let i = 0; i < lines.length; i++) {
            const trimmedLine = lines[i].trim();
            
            // 跳过标题行和空行
            if (trimmedLine.includes('Pin#') || trimmedLine.includes('Pin Name') || 
                trimmedLine.includes('Description') || trimmedLine.length < 3) {
                continue;
            }
            
            // 匹配行首的数字（可能包含逗号分隔的多个数字）
            const pinNumberMatch = trimmedLine.match(/^(\d+(?:\s*,\s*\d+)*)\s+/);
            if (!pinNumberMatch) {
                continue;
            }
            
            const pinNumberStr = pinNumberMatch[1];
            const restOfLine = trimmedLine.substring(pinNumberMatch[0].length).trim();
            
            console.log(`行 ${i + 1}: 引脚编号="${pinNumberStr}", 剩余文本="${restOfLine}"`);
            
            // 解析引脚编号（处理多个编号的情况，如 "16, 25, 62, 7"）
            const pinNumbers = pinNumberStr.split(',').map(n => parseInt(n.trim())).filter(n => !isNaN(n) && n > 0);
            
            if (pinNumbers.length === 0) {
                continue;
            }
            
            // 提取引脚名称（通常是第一个字母数字组合）
            const pinNameMatch = restOfLine.match(/^([A-Za-z][A-Za-z0-9_\-\/\.\+\*]*)/);
            if (!pinNameMatch) {
                console.log(`  -> 跳过：未找到有效的引脚名称`);
                continue;
            }
            
            const pinName = pinNameMatch[1].trim();
            
            // 验证引脚名称
            if (!isValidPinName(pinName)) {
                console.log(`  -> 跳过：无效的引脚名称 "${pinName}"`);
                continue;
            }
            
            console.log(`  -> 引脚名称="${pinName}", 编号数量=${pinNumbers.length}`);
            
            // 为每个引脚编号创建记录
            for (const pinNumber of pinNumbers) {
                if (pinNumber > 0 && pinNumber <= 1000 && !seenNumbers.has(pinNumber)) {
                    seenNumbers.add(pinNumber);
                    
                    pins.push({
                        Number: pinNumber,
                        Name: pinName,
                        Type: 'Passive', // 默认值设为 Passive
                        Visible: 'TRUE',
                        Shape: 'Line',
                        PinGroup: '', // 保持空白
                        Position: '', // 由自动排列功能分配
                        Section: 'A'
                    });
                }
            }
        }
    }
    
    console.log(`提取完成！共识别到 ${pins.length} 个引脚`);
    
    // 按引脚编号排序
    pins.sort((a, b) => a.Number - b.Number);
    
    // 计算总引脚数（最大编号）
    let totalPins = 0;
    if (pins.length > 0) {
        totalPins = Math.max(...pins.map(pin => pin.Number));
    }
    
    return {
        totalPins,
        pins
    };
}

// 验证引脚名称是否有效
function isValidPinName(name) {
    if (!name || name.length === 0) return false;
    
    // 排除常见的非引脚词汇
    const invalidWords = [
        'pin', 'pins', 'description', 'table', 'figure', 'fig', 
        'page', 'note', 'notes', 'signal', 'type', 'direction',
        'level', 'voltage', 'current'
    ];
    
    const lowerName = name.toLowerCase();
    if (invalidWords.includes(lowerName)) return false;
    
    // 引脚名称应该包含字母或数字
    if (!/[A-Za-z0-9]/.test(name)) return false;
    
    // 引脚名称长度应该在合理范围内
    if (name.length > 50) return false;
    
    return true;
}

// 从粘贴的文本中提取引脚信息
function extractPinDataFromPastedText(text) {
    const pins = [];
    const seenNumbers = new Set();
    const lines = text.split('\n').filter(line => line.trim().length > 0);
    
    console.log('从粘贴文本中提取引脚信息...');
    console.log('总行数:', lines.length);
    
    let pendingLine = ''; // 用于处理跨行的 Pin Name
    
    for (let i = 0; i < lines.length; i++) {
        let trimmedLine = lines[i].trim();
        
        // 如果是延续上一行的内容
        if (pendingLine) {
            trimmedLine = pendingLine + ' ' + trimmedLine;
            pendingLine = '';
        }
        
        if (trimmedLine.length < 2) {
            continue;
        }
        
        // 分割行为多个部分
        const parts = trimmedLine.split(/\s+/).filter(p => p.length > 0);
        
        if (parts.length >= 2) {
            let pinNumber = '';
            let pinName = '';
            
            // 尝试识别哪个是 Pin #，哪个是 Pin Name
            const firstPart = parts[0];
            const secondPart = parts[1];
            
            // 检查第一个部分是否包含数字
            const firstHasDigit = /\d/.test(firstPart);
            const secondHasDigit = /\d/.test(secondPart);
            
            // 检查是否以斜杠开头（表示是上一行的延续）
            const startsWithSlash = firstPart.startsWith('/');
            
            if (startsWithSlash && parts.length >= 3) {
                // 第一个部分以/开头，说明这是上一行 Pin Name 的延续
                // 需要找到 Pin #（应该在后面的部分中）
                for (let j = 1; j < parts.length; j++) {
                    if (/\d/.test(parts[j])) {
                        pinNumber = parts[j];
                        // Pin Name 是除了 Pin # 之外的所有部分
                        pinName = parts.slice(0, j).join('') + parts.slice(j + 1).join('');
                        break;
                    }
                }
            } else if (firstHasDigit && !secondHasDigit) {
                // 第一个是 Pin #（包含数字），第二个是 Pin Name
                pinNumber = firstPart;
                pinName = parts.slice(1).join(' ');
            } else if (!firstHasDigit && secondHasDigit) {
                // 第一个是 Pin Name，第二个是 Pin #
                pinName = firstPart;
                pinNumber = secondPart;
            } else if (firstHasDigit && secondHasDigit) {
                // 都包含数字，假设第一个是 Pin #
                pinNumber = firstPart;
                pinName = parts.slice(1).join(' ');
            }
            
            // 验证：如果 Pin Name 看起来不完整（以斜杠结尾，或包含未闭合的方括号）
            if (pinName && (pinName.endsWith('/') || (pinName.includes('[') && !pinName.includes(']')))) {
                // 保存到待处理行，等待下一行继续
                pendingLine = trimmedLine;
                continue;
            }
            
            // 验证引脚名称
            if (pinName && pinNumber && !seenNumbers.has(pinNumber)) {
                seenNumbers.add(pinNumber);
                
                console.log(`行 ${i + 1}: Pin #="${pinNumber}", Pin Name="${pinName}"`);
                
                pins.push({
                    Number: pinNumber,
                    Name: pinName,
                    Type: 'Passive',
                    Visible: 'TRUE',
                    Shape: 'Line',
                    PinGroup: '',
                    Position: '',
                    Section: 'A'
                });
            }
        }
    }
    
    console.log(`从粘贴文本中提取到 ${pins.length} 个引脚`);
    return pins;
}

// 智能排序函数（支持字母数字混合）
function smartSortPins() {
    if (pinData.length === 0) {
        alert('没有数据可排序');
        return;
    }
    
    // 自然排序（支持字母数字混合，如 AJ26, AK2, B10 等）
    pinData.sort((a, b) => {
        const aNum = a.Number.toString();
        const bNum = b.Number.toString();
        
        // 使用 localeCompare 进行自然排序
        return aNum.localeCompare(bNum, 'en', {
            numeric: true,
            sensitivity: 'base'
        });
    });
    
    displayPinData();
    alert('数据已按自然顺序排序！');
}

// 提取引脚信息函数
async function extractPinInfo() {
    const fileInput = document.getElementById('fileInput');
    const files = fileInput.files;
    const pastedPinData = document.getElementById('pastedPinData')?.value.trim() || '';
    const pdfPageRange = document.getElementById('pdfPageRange').value;
    
    if (files.length === 0 && pastedPinData === '') {
        alert('请先选择文件或粘贴引脚数据');
        return;
    }
    
    // 如果用户粘贴了数据，优先使用粘贴的数据
    if (pastedPinData) {
        try {
            const pastedPins = extractPinDataFromPastedText(pastedPinData);
            
            if (pastedPins.length === 0) {
                alert('未能从粘贴的文本中识别到引脚数据，请检查格式');
                return;
            }
            
            // 更新芯片信息
            const maxPinNum = Math.max(...pastedPins.map(p => {
                const match = p.Number.match(/\d+/);
                return match ? parseInt(match[0]) : 0;
            }));
            chipInfo.totalPins = maxPinNum || pastedPins.length;
            
            // 更新引脚数据
            pinData = pastedPins;
            
            // 自动应用 Position 分配
            const autoArrange = document.getElementById('autoArrange')?.checked !== false;
            if (autoArrange && pinData.length > 0) {
                reapplyPositionAllocation();
            } else {
                displayPinData();
            }
            
            alert(`成功识别 ${pastedPins.length} 个引脚！`);
            return;
        } catch (error) {
            console.error('解析粘贴数据失败:', error);
            alert('解析粘贴数据失败，请检查格式');
        }
    }
    
    try {
        // 显示进度区域
        const progressSection = document.getElementById('progressSection');
        progressSection.style.display = 'block';
        
        let extractedText = '';
        
        // 处理每个文件
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const fileExtension = file.name.split('.').pop().toLowerCase();
            
            if (fileExtension === 'pdf') {
                // 处理 PDF 文件
                document.getElementById('pdfProgress').textContent = `正在处理 PDF 文件 (${i + 1}/${files.length})...`;
                const pageRange = parsePageRange(pdfPageRange);
                const pdfText = await extractTextFromPDF(file, pageRange);
                extractedText += pdfText + ' ';
            } else if (['jpg', 'jpeg', 'png'].includes(fileExtension)) {
                // 处理图片文件
                document.getElementById('ocrProgress').textContent = `正在识别图片 (${i + 1}/${files.length})...`;
                const imageText = await extractTextFromImage(URL.createObjectURL(file));
                extractedText += imageText + ' ';
            }
        }
        
        // 显示提取的原始文本（用于调试）
        document.getElementById('extractedText').value = extractedText;
        
        // 从提取的文本中解析引脚信息
        const extractionResult = extractPinInfoFromText(extractedText);
        
        // 更新芯片信息
        chipInfo.totalPins = extractionResult.totalPins;
        
        // 更新 UI 显示
        document.getElementById('totalPinsDisplay').textContent = extractionResult.totalPins > 0 ? extractionResult.totalPins : '未识别';
        document.getElementById('totalPins').value = extractionResult.totalPins;
        
        // 更新引脚数据
        pinData = extractionResult.pins;
        
        // 更新 chipInfo.totalPins
        chipInfo.totalPins = extractionResult.totalPins;
        
        // 自动应用 Position 分配（如果启用了自动排列）
        const autoArrange = document.getElementById('autoArrange')?.checked !== false;
        if (autoArrange && pinData.length > 0) {
            reapplyPositionAllocation();
        } else {
            // 显示提取的数据
            displayPinData();
        }
        
        // 隐藏进度区域
        progressSection.style.display = 'none';
        document.getElementById('pdfProgress').textContent = '';
        document.getElementById('ocrProgress').textContent = '';
        
        alert('引脚信息提取完成！');
    } catch (error) {
        console.error('提取信息时出错:', error);
        document.getElementById('progressSection').style.display = 'none';
        alert('提取信息时出错，请检查控制台获取详细信息');
    }
}

// 重新应用 Position 分配
function reapplyPositionAllocation() {
    if (pinData.length === 0) return;
    
    const autoArrange = document.getElementById('autoArrange')?.checked !== false;
    const reverseBottomTop = document.getElementById('reverseBottomTop')?.checked || false;
    const totalPins = chipInfo.totalPins || Math.max(...pinData.map(p => p.Number));
    const quarter = Math.ceil(totalPins / 4);
    
    pinData.forEach(pin => {
        const pinNumber = parseInt(pin.Number);
        
        if (autoArrange && pinNumber > 0) {
            const pinInPackage = ((pinNumber - 1) % totalPins + totalPins) % totalPins + 1;
            
            if (pinInPackage >= 1 && pinInPackage <= quarter) {
                pin.Position = 'LEFT';
            } else if (pinInPackage > quarter && pinInPackage <= quarter * 2) {
                pin.Position = reverseBottomTop ? 'BOTTOM_REV' : 'BOTTOM';
            } else if (pinInPackage > quarter * 2 && pinInPackage <= quarter * 3) {
                pin.Position = 'RIGHT';
            } else if (pinInPackage > quarter * 3 && pinInPackage <= totalPins) {
                pin.Position = reverseBottomTop ? 'TOP_REV' : 'TOP';
            } else {
                pin.Position = '';
            }
        } else {
            // 手动模式，清空 Position
            pin.Position = '';
        }
    });
    
    // 重新显示数据
    displayPinData();
}

// 显示引脚数据
function displayPinData() {
    const tbody = document.querySelector('#pinTable tbody');
    tbody.innerHTML = '';
    
    pinData.forEach(pin => {
        const row = document.createElement('tr');
        const keys = Object.keys(pin);
        
        // 创建可编辑的单元格
        keys.forEach((key, index) => {
            const cell = document.createElement('td');
            const value = pin[key];
            
            let input;
            
            // Position 列使用下拉选择框
            if (key === 'Position') {
                input = document.createElement('select');
                const options = ['', 'TOP', 'TOP_REV', 'BOTTOM', 'BOTTOM_REV', 'LEFT', 'RIGHT'];
                const optionLabels = {
                    '': '',
                    'TOP': 'TOP (升序)',
                    'TOP_REV': 'TOP (降序)',
                    'BOTTOM': 'BOTTOM (升序)',
                    'BOTTOM_REV': 'BOTTOM (降序)',
                    'LEFT': 'LEFT',
                    'RIGHT': 'RIGHT'
                };
                options.forEach(opt => {
                    const option = document.createElement('option');
                    option.value = opt;
                    option.textContent = optionLabels[opt] || opt;
                    if (value === opt) {
                        option.selected = true;
                    }
                    input.appendChild(option);
                });
            } else {
                // 其他列使用文本输入框
                input = document.createElement('input');
                input.type = 'text';
                input.value = value;
            }
            
            input.addEventListener('change', function() {
                // 更新数据
                const rowIndex = Array.from(tbody.children).indexOf(row);
                const dataKey = keys[index];
                pinData[rowIndex][dataKey] = this.value;
            });
            
            cell.appendChild(input);
            row.appendChild(cell);
        });
        
        tbody.appendChild(row);
    });
}



// 导出到 Excel
function exportToExcel() {
    // 检查是否有数据
    if (pinData.length === 0) {
        alert('没有数据可导出');
        return;
    }
    
    // 定义列顺序（按照示例图片的格式）
    const columnOrder = ['Number', 'Name', 'Type', 'Visible', 'Shape', 'PinGroup', 'Position', 'Section'];
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建工作表数据 - 表头
    const wsData = [columnOrder];
    
    // 添加数据行
    pinData.forEach(pin => {
        const row = columnOrder.map(col => {
            let value = pin[col];
            // 确保 Number 列是数字类型
            if (col === 'Number') {
                return parseInt(value) || 0;
            }
            // Position 列：将 TOP_REV/BOTTOM_REV 转换为 TOP/BOTTOM
            if (col === 'Position') {
                if (value === 'TOP_REV') {
                    value = 'TOP';
                } else if (value === 'BOTTOM_REV') {
                    value = 'BOTTOM';
                }
            }
            return value || '';
        });
        wsData.push(row);
    });
    
    // 创建工作表
    const ws = XLSX.utils.aoa_to_sheet(wsData, {
        skipHeader: false
    });
    
    // 设置列宽
    const colWidths = [
        { wch: 8 },  // Number
        { wch: 20 }, // Name
        { wch: 12 }, // Type
        { wch: 8 },  // Visible
        { wch: 8 },  // Shape
        { wch: 12 }, // PinGroup
        { wch: 10 }, // Position
        { wch: 8 }   // Section
    ];
    ws['!cols'] = colWidths;
    
    // 设置表头样式（加粗）
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_col(C) + "1";
        if (!ws[address]) continue;
        ws[address].s = {
            font: { bold: true }
        };
    }
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, 'PinInfo');
    
    // 导出 Excel 文件
    const filename = `chip_pinout_${new Date().getTime()}.xlsx`;
    XLSX.writeFile(wb, filename);
    alert(`Excel 文件已导出：${filename}`);
}

// 加载SheetJS库
(function() {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
    script.onload = function() {
        console.log('SheetJS库加载完成');
    };
    document.head.appendChild(script);
})();

// 加载Tesseract.js库用于OCR
(function() {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@2.1.5/dist/tesseract.min.js';
    script.onload = function() {
        console.log('Tesseract.js库加载完成');
    };
    document.head.appendChild(script);
})();

// 加载PDF.js库用于处理PDF文件
(function() {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.10.377/build/pdf.min.js';
    script.onload = function() {
        console.log('PDF.js库加载完成');
    };
    document.head.appendChild(script);
})();