// 全局变量定义
let baseFile = null;
let compareFile = null;
let baseFileContent = null;
let compareFileContent = null;
let baseFileMarked = false;
let compareFileMarked = false;
let diffResult = null;

// 库加载状态检查
let diffLibraryReady = false;

// 检查Diff库是否加载
function checkDiffLibrary() {
    try {
        if (typeof Diff !== 'undefined' && typeof Diff.diffArrays === 'function') {
            diffLibraryReady = true;
            console.log('✅ Diff库加载成功，可以使用');
            return true;
        } else {
            diffLibraryReady = false;
            console.error('❌ Diff库未正确加载，Diff对象或方法不存在');
            return false;
        }
    } catch (error) {
        diffLibraryReady = false;
        console.error('❌ 检查Diff库时出错:', error);
        return false;
    }
}

// DOM元素获取
const baseFileInput = document.getElementById('baseFileInput');
const compareFileInput = document.getElementById('compareFileInput');
const baseFileDrop = document.getElementById('baseFileDrop');
const compareFileDrop = document.getElementById('compareFileDrop');
const baseFileName = document.getElementById('baseFileName');
const compareFileName = document.getElementById('compareFileName');
const compareButton = document.getElementById('compareButton');
const markBaseButton = document.getElementById('markBaseChanges');
const markCompareButton = document.getElementById('markCompareChanges');
const downloadBaseDocxButton = document.getElementById('downloadBaseDocx');
const downloadCompareDocxButton = document.getElementById('downloadCompareDocx');
const downloadBasePdfButton = document.getElementById('downloadBasePdf');
const downloadComparePdfButton = document.getElementById('downloadComparePdf');
const baseFileDisplay = document.getElementById('baseFileDisplay');
const compareFileDisplay = document.getElementById('compareFileDisplay');
const baseFileStatus = document.getElementById('baseFileStatus');
const compareFileStatus = document.getElementById('compareFileStatus');
const resultSection = document.getElementById('resultSection');
const statsSection = document.getElementById('statsSection');
const totalChanges = document.getElementById('totalChanges');
const addedWords = document.getElementById('addedWords');
const deletedWords = document.getElementById('deletedWords');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingText = document.getElementById('loadingText');
const baseFileInfo = document.getElementById('baseFileInfo');
const compareFileInfo = document.getElementById('compareFileInfo');
const removeBaseFile = document.getElementById('removeBaseFile');
const removeCompareFile = document.getElementById('removeCompareFile');

// 原始内容和标记内容容器
const baseFileRawContent = document.querySelector('#baseFileRawContent > div');
const baseFileMarkedContent = document.querySelector('#baseFileMarkedContent > div');
const compareFileRawContent = document.querySelector('#compareFileRawContent > div');
const compareFileMarkedContent = document.querySelector('#compareFileMarkedContent > div');

// 初始化事件监听器
function initEventListeners() {
    // 文件上传拖放功能
    setupFileDrop(baseFileDrop, baseFileInput, handleBaseFileSelect);
    setupFileDrop(compareFileDrop, compareFileInput, handleCompareFileSelect);
    
    // 添加文件上传动画效果
    baseFileDrop.addEventListener('dragover', () => {
        baseFileDrop.classList.add('scale-105');
    });
    
    baseFileDrop.addEventListener('dragleave', () => {
        baseFileDrop.classList.remove('scale-105');
    });
    
    compareFileDrop.addEventListener('dragover', () => {
        compareFileDrop.classList.add('scale-105');
    });
    
    compareFileDrop.addEventListener('dragleave', () => {
        compareFileDrop.classList.remove('scale-105');
    });
    
    // 文件删除按钮
    removeBaseFile.addEventListener('click', clearBaseFile);
    removeCompareFile.addEventListener('click', clearCompareFile);
    
    // 操作按钮
    compareButton.addEventListener('click', compareFiles);
    markBaseButton.addEventListener('click', markBaseFileChanges);
    markCompareButton.addEventListener('click', markCompareFileChanges);
    
    // 下载按钮事件处理已在index.html中定义，这里不再重复绑定
}

// 设置文件拖放功能
function setupFileDrop(dropZone, inputField, handleFileSelect) {
    // 点击区域触发文件选择
    dropZone.addEventListener('click', () => inputField.click());
    
    // 文件选择事件
    inputField.addEventListener('change', async (e) => {
        if (e.target.files.length > 0) {
            await handleFileSelect(e.target.files[0]);
        }
    });
    
    // 拖放事件
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, unhighlight, false);
    });
    
    function highlight() {
        dropZone.classList.add('file-drop-active');
    }
    
    function unhighlight() {
        dropZone.classList.remove('file-drop-active');
    }
    
    // 处理文件拖放
    dropZone.addEventListener('drop', async (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        
        if (files.length > 0) {
            await handleFileSelect(files[0]);
        }
    });
}

// 处理基准文件选择
async function handleBaseFileSelect(file) {
    if (!file.name.endsWith('.docx')) {
        alert('请选择 .docx 格式的文件');
        return;
    }
    
    baseFile = file;
    baseFileName.textContent = file.name;
    baseFileInfo.classList.remove('hidden');
    
    // 显示加载状态
    baseFileStatus.textContent = '加载中...';
    baseFileStatus.classList.add('loading');
    
    try {
        // 解析文件内容
        baseFileContent = await parseDocxFile(file);
        baseFileStatus.textContent = '已加载';
        baseFileStatus.classList.remove('loading');
        baseFileStatus.classList.add('success');
    } catch (error) {
        console.error('文件解析错误:', error);
        baseFileStatus.textContent = '加载失败';
        baseFileStatus.classList.remove('loading');
        baseFileStatus.classList.add('error');
        alert('文件解析失败: ' + error.message);
        return;
    }
    
    updateButtonStates();
}

// 处理对比文件选择
async function handleCompareFileSelect(file) {
    if (!file.name.endsWith('.docx')) {
        alert('请选择 .docx 格式的文件');
        return;
    }
    
    compareFile = file;
    compareFileName.textContent = file.name;
    compareFileInfo.classList.remove('hidden');
    
    // 显示加载状态
    compareFileStatus.textContent = '加载中...';
    compareFileStatus.classList.add('loading');
    
    try {
        // 解析文件内容
        compareFileContent = await parseDocxFile(file);
        compareFileStatus.textContent = '已加载';
        compareFileStatus.classList.remove('loading');
        compareFileStatus.classList.add('success');
    } catch (error) {
        console.error('文件解析错误:', error);
        compareFileStatus.textContent = '加载失败';
        compareFileStatus.classList.remove('loading');
        compareFileStatus.classList.add('error');
        alert('文件解析失败: ' + error.message);
        return;
    }
    
    updateButtonStates();
}

// 清除基准文件
function clearBaseFile() {
    baseFile = null;
    baseFileContent = null;
    baseFileInput.value = '';
    baseFileName.textContent = '';
    baseFileStatus.textContent = '';
    baseFileStatus.classList.remove('success', 'error', 'loading');
    baseFileInfo.classList.add('hidden');
    resetResults();
    updateButtonStates();
}

// 清除对比文件
function clearCompareFile() {
    compareFile = null;
    compareFileContent = null;
    compareFileInput.value = '';
    compareFileName.textContent = '';
    compareFileStatus.textContent = '';
    compareFileStatus.classList.remove('success', 'error', 'loading');
    compareFileInfo.classList.add('hidden');
    resetResults();
    updateButtonStates();
}

// 更新按钮状态
function updateButtonStates() {
    const filesReady = baseFile && compareFile;
    compareButton.disabled = !filesReady;
    
    const comparisonDone = diffResult !== null;
    markBaseButton.disabled = !comparisonDone;
    markCompareButton.disabled = !comparisonDone;
    
    // 根据文件状态启用下载按钮
    const canDownloadBase = comparisonDone && (baseFileMarked || baseFileContent);
    const canDownloadCompare = comparisonDone && (compareFileMarked || compareFileContent);
    
    downloadBaseDocxButton.disabled = !canDownloadBase;
    downloadBasePdfButton.disabled = !canDownloadBase;
    downloadCompareDocxButton.disabled = !canDownloadCompare;
    downloadComparePdfButton.disabled = !canDownloadCompare;
}

// 重置结果显示
function resetResults() {
    resultSection.classList.add('hidden');
    statsSection.classList.add('hidden');
    diffResult = null;
    baseFileMarked = false;
    compareFileMarked = false;
    baseFileStatus.textContent = '未标记修改';
    compareFileStatus.textContent = '未标记修改';
    
    // 清空原始内容和标记内容
    baseFileRawContent.innerHTML = '';
    baseFileMarkedContent.innerHTML = '';
    compareFileRawContent.innerHTML = '';
    compareFileMarkedContent.innerHTML = '';
    
    // 清空统计信息
    totalChanges.textContent = '0';
    addedWords.textContent = '0';
    deletedWords.textContent = '0';
    
    updateButtonStates();
}

// 显示加载状态
function showLoading(text = '处理中，请稍候...') {
    loadingText.textContent = text;
    loadingOverlay.classList.remove('hidden');
}

// 隐藏加载状态
function hideLoading() {
    loadingOverlay.classList.add('hidden');
}

// 解析docx文件
function parseDocxFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function(e) {
            const arrayBuffer = e.target.result;

            // 配置mammoth以保留表格格式
            const mammothOptions = {
                arrayBuffer: arrayBuffer,
                // 保留表格结构
                includeDefaultStyleMap: true,
                // 自定义样式映射以保留表格
                styleMap: [
                    "p[style-name='Heading 1] => h1:fresh",
                    "p[style-name='Heading 2] => h2:fresh",
                    "p[style-name='Heading 3] => h3:fresh",
                    "p[style-name='Heading 4] => h4:fresh",
                    "p[style-name='Heading 5] => h5:fresh",
                    "p[style-name='Heading 6] => h6:fresh",
                    "p[style-name='Title'] => h1.title:fresh",
                    "p[style-name='Subtitle'] => h2.subtitle:fresh",
                    // 保留表格格式
                    "table => table",
                    "tr => tr",
                    "td => td",
                    "th => th"
                ].join("\n"),
                // 转换表格为HTML表格元素
                convertMarkdown: false
            };

            // 并行提取纯文本和HTML格式
            Promise.all([
                mammoth.extractRawText({ arrayBuffer: arrayBuffer }),
                mammoth.convertToHtml(mammothOptions)
            ])
            .then(([textResult, htmlResult]) => {
                // 后处理HTML以确保表格结构完整
                let processedHtml = htmlResult.value;

                // 确保表格有完整的结构
                processedHtml = processedHtml.replace(/<table([^>]*)>/gi, '<table$1 border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse;">');

                // 确保表格单元格有适当的样式
                processedHtml = processedHtml.replace(/<td([^>]*)>/gi, '<td$1 style="border: 1px solid #ddd; padding: 8px;">');
                processedHtml = processedHtml.replace(/<th([^>]*)>/gi, '<th$1 style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5;">');

                resolve({
                    text: textResult.value,
                    html: processedHtml,
                    rawHtml: processedHtml // 保存处理后的HTML
                });
            })
            .catch(error => {
                console.error('文档解析错误:', error);
                reject(new Error('文档解析失败: ' + (error.message || '未知错误')));
            });
        };

        reader.onerror = function() {
            reject(new Error('文件读取失败'));
        };

        reader.readAsArrayBuffer(file);
    });
}

// 按照要求进行分词
function tokenizeText(text) {
    if (!text) return [];
    
    const tokens = [];
    let i = 0;
    
    // 按照要求的分词逻辑处理：英文单词、数字、中文、空格、标点符号
    while (i < text.length) {
        const char = text[i];
        
        // 1. 处理英文单词（连续字母）
        if (/[a-zA-Z]/.test(char)) {
            let word = '';
            while (i < text.length && /[a-zA-Z]/.test(text[i])) {
                word += text[i];
                i++;
            }
            tokens.push(word);
            continue;
        }
        
        // 2. 处理数字（连续数字）
        if (/[0-9]/.test(char)) {
            let number = '';
            while (i < text.length && /[0-9]/.test(text[i])) {
                number += text[i];
                i++;
            }
            tokens.push(number);
            continue;
        }
        
        // 3. 处理空格（单独作为一个处理单位）
        if (char === ' ') {
            tokens.push(char);
            i++;
            continue;
        }
        
        // 4. 处理中文或其他字符（每个汉字或符号单独处理）
        tokens.push(char);
        i++;
    }
    
    return tokens;
}

// 对比文件
async function compareFiles() {
    showLoading('正在对比文件...');
    
    try {
        // 检查文件内容是否已加载
        if (!baseFileContent) {
            throw new Error('基准文件内容未加载，请重新选择基准文件');
        }
        if (!compareFileContent) {
            throw new Error('对比文件内容未加载，请重新选择对比文件');
        }
        
        // 分词处理 - 将文本分割为单词、数字、空格和其他字符
        const baseTokens = tokenizeText(baseFileContent.text);
        const compareTokens = tokenizeText(compareFileContent.text);
        
        // 使用token数组进行单词级别的比较
    // 注意：我们使用tokens数组而不是原始文本进行比较，确保英文单词级别的差异识别
    if (!checkDiffLibrary()) {
        showLoading('差异比较库加载失败，无法继续比较操作', true);
        setTimeout(hideLoading, 3000);
        throw new Error('Diff库未加载');
    }
    
    const tokenDiffResult = Diff.diffArrays(baseTokens, compareTokens);

        // 将token级别的差异转换回文本级别的差异，方便显示
        const textDiff = [];
        tokenDiffResult.forEach(part => {
            textDiff.push({
                added: part.added,
                removed: part.removed,
                value: Array.isArray(part.value) ? part.value.join('') : part.value
            });
        });

        // 优化差异结果
        const optimizedDiff = optimizeDiffResult(textDiff);
        
        // 存储结果
        const finalDiffResult = {
            diff: optimizedDiff,
            baseContent: baseFileContent.text,
            compareContent: compareFileContent.text,
            baseTokens: baseTokens,
            compareTokens: compareTokens
        };

        // 将结果赋值给全局变量
        diffResult = finalDiffResult;
        
        // 显示结果
        displayResults();

        // 计算并更新统计信息
        updateStats(optimizedDiff);
        
        // 显示结果区域，先检查元素是否存在
        if (resultSection) {
            resultSection.classList.remove('hidden');
        } else {
            console.error('Error: resultSection element not found');
        }
        
        // 更新按钮状态
        updateButtonStates();
    } catch (error) {
        console.error('文件对比失败:', error);
        alert('文件对比失败: ' + error.message);
    } finally {
        hideLoading();
    }
}

// 优化差异结果，避免整段标记
function optimizeDiffResult(diff) {
    const optimized = [];
    const MAX_CONTINUOUS_CHANGES = 10; // 连续变化的最大数量
    
    let i = 0;
    while (i < diff.length) {
        // 检查是否有连续的大量变化
        if (isSignificantChange(diff[i]) && i + 1 < diff.length) {
            let continuousChanges = 1;
            let j = i + 1;
            
            // 计算连续的变化数量
            while (j < diff.length && isSignificantChange(diff[j])) {
                continuousChanges++;
                j++;
            }
            
            // 如果连续变化数量超过阈值，尝试保留一些上下文
            if (continuousChanges > MAX_CONTINUOUS_CHANGES) {
                // 保留变化部分，但确保标记更精确
                for (let k = i; k < j; k++) {
                    // 对于长段落，我们仍然标记变化，但确保分词正确
                    optimized.push(diff[k]);
                }
                i = j;
                continue;
            }
        }
        
        // 正常添加
        optimized.push(diff[i]);
        i++;
    }
    
    return optimized;
}

// 判断是否为重要变化
function isSignificantChange(part) {
    return part.added || part.removed;
}

// 计算差异统计信息
function calculateDiffStats(diff) {
    let modifications = 0;
    let additions = 0;
    let deletions = 0;
    let equalCount = 0;
    let totalCount = 0;
    
    diff.forEach(part => {
        if (part.added) {
            additions++;
        } else if (part.removed) {
            deletions++;
        } else {
            equalCount++;
        }
        totalCount++;
    });
    
    // 计算修改数量（假设删除和新增相邻的部分为修改）
    for (let i = 0; i < diff.length - 1; i++) {
        if (diff[i].removed && diff[i + 1].added) {
            modifications++;
        }
    }
    
    // 计算相似度
    const similarity = totalCount > 0 ? (equalCount / totalCount * 100) : 100;
    
    return {
        modifications,
        additions,
        deletions,
        similarity
    };
}

// 显示对比结果
function displayResults() {
    console.log('displayResults called:', {
        diffResult: !!diffResult,
        baseFileContent: !!baseFileContent,
        compareFileContent: !!compareFileContent,
        baseFileRawContent: !!baseFileRawContent,
        compareFileRawContent: !!compareFileRawContent
    });

    if (!diffResult || !baseFileContent || !compareFileContent) {
        console.error('Missing data for displayResults');
        return;
    }

    // 显示原始内容，保留表格结构和换行符，先检查元素是否存在
    if (baseFileRawContent) {
        // 使用HTML内容而不是纯文本，以保留表格结构
        baseFileRawContent.innerHTML = baseFileContent.html || escapeHtml(baseFileContent.text).replace(/\n/g, '<br>');
    } else {
        console.error('Error: baseFileRawContent element not found');
    }

    if (compareFileRawContent) {
        // 使用HTML内容而不是纯文本，以保留表格结构
        compareFileRawContent.innerHTML = compareFileContent.html || escapeHtml(compareFileContent.text).replace(/\n/g, '<br>');
    } else {
        console.error('Error: compareFileRawContent element not found');
    }

    // 初始清空标记内容，先检查元素是否存在
    if (baseFileMarkedContent) {
        baseFileMarkedContent.innerHTML = '<p class="text-gray-500 italic">点击"标记基准文件修改"按钮查看标记内容</p>';
    } else {
        console.error('Error: baseFileMarkedContent element not found');
    }

    if (compareFileMarkedContent) {
        compareFileMarkedContent.innerHTML = '<p class="text-gray-500 italic">点击"标记对比文件修改"按钮查看标记内容</p>';
    } else {
        console.error('Error: compareFileMarkedContent element not found');
    }

    // 重置状态显示，先检查元素是否存在
    if (baseFileStatus) {
        baseFileStatus.textContent = '未标记';
        baseFileStatus.classList.remove('text-primary', 'font-medium');
    } else {
        console.error('Error: baseFileStatus element not found');
    }

    if (compareFileStatus) {
        compareFileStatus.textContent = '未标记';
        compareFileStatus.classList.remove('text-accent', 'font-medium');
    } else {
        console.error('Error: compareFileStatus element not found');
    }

    // 重置标记状态
    baseFileMarked = false;
    compareFileMarked = false;

    // 显示结果区域，先检查元素是否存在
    if (resultSection) {
        resultSection.classList.remove('hidden');
    } else {
        console.error('Error: resultSection element not found');
    }

    if (statsSection) {
        statsSection.classList.remove('hidden');
    } else {
        console.error('Error: statsSection element not found');
    }
}

// 更新统计信息
function updateStats(diff) {
    if (!diff || !diffResult) return;
    
    // 计算修改数量、添加字数和删除字数
    let changesCount = 0;
    let addedCount = 0;
    let removedCount = 0;
    
    // 遍历差异结果统计信息
    diff.forEach(part => {
        if (part.added) {
            addedCount += part.value.length;
            changesCount++;
        } else if (part.removed) {
            removedCount += part.value.length;
            changesCount++;
        }
    });
    
    // 更新统计信息，先检查元素是否存在
    if (totalChanges) {
        totalChanges.textContent = changesCount;
    } else {
        console.error('Error: totalChanges element not found');
    }
    
    if (addedWords) {
        addedWords.textContent = addedCount;
    } else {
        console.error('Error: addedWords element not found');
    }
    
    if (deletedWords) {
        deletedWords.textContent = removedCount;
    } else {
        console.error('Error: deletedWords element not found');
    }
    
    // 计算相似度百分比
    const totalChars = diffResult.baseContent.length + diffResult.compareContent.length;
    const similarity = totalChars > 0 ? 
        ((totalChars - addedCount - removedCount) / totalChars * 100).toFixed(2) : 
        '100.00';
    
    // 添加相似度显示
    const similarityText = document.createElement('div');
    similarityText.id = 'similarityText';
    similarityText.textContent = `相似度: ${similarity}%`;
    
    // 检查是否已存在相似度元素，如果存在则更新，否则添加
    const existingSimilarity = document.getElementById('similarityText');
    if (existingSimilarity) {
        existingSimilarity.textContent = `相似度: ${similarity}%`;
    } else if (statsSection) {
        statsSection.appendChild(similarityText);
    } else {
        console.error('Error: statsSection element not found for appending similarityText');
    }
    
    // 显示统计信息区域，先检查元素是否存在
    if (statsSection) {
        statsSection.classList.remove('hidden');
    } else {
        console.error('Error: statsSection element not found');
    }
}

// 标记基准文件修改
async function markBaseFileChanges() {
    console.log('markBaseFileChanges called:', {
        diffResult: !!diffResult,
        diffResultDiff: !!(diffResult && diffResult.diff),
        baseContent: !!(diffResult && diffResult.baseContent),
        baseFileMarkedContent: !!baseFileMarkedContent
    });

    if (!diffResult || !diffResult.diff || !diffResult.baseContent) return;

    // 显示加载状态
    markBaseButton.disabled = true;
    showLoading('正在标记基准文件...');

    try {
        // 重置标记状态
        baseFileMarked = false;

        // 创建标记内容 - 基于HTML内容进行标记
        let markedHtml = baseFileContent.html || '';

        // 简化版本：直接使用文本标记方法，但包装在HTML标签中
        let markedContent = '';
        diffResult.diff.forEach(part => {
            if (part.removed) {
                markedContent += `<span class="modified-text">${escapeHtml(part.value)}</span>`;
            } else if (!part.added) {
                markedContent += escapeHtml(part.value);
            }
        });
        // 包装在段落标签中，确保DOCX生成能正确解析
        markedHtml = `<p>${markedContent.replace(/\n\n+/g, '</p><p>').replace(/\n/g, '<br>')}</p>`;

        // 显示标记内容
        console.log('Setting base marked content:', markedHtml.substring(0, 500) + '...');
        console.log('Content contains modified-text:', markedHtml.includes('modified-text'));
        console.log('Content contains added-text:', markedHtml.includes('added-text'));
        baseFileMarkedContent.innerHTML = markedHtml;

        // 保存标记内容到baseFileContent对象
        baseFileContent.markedHtml = markedHtml;
        baseFileContent.markedText = markedHtml;

        // 更新状态
        baseFileMarked = true;
        baseFileStatus.textContent = '已标记';
        baseFileStatus.classList.add('text-primary', 'font-medium');

        // 滚动到结果区域
        baseFileMarkedContent.scrollIntoView({ behavior: 'smooth', block: 'start' });

        // 更新按钮状态
        updateButtonStates();

        // 保存内容以供下载
        saveContentForDownload('base', '基准文件');

        console.log('Base file marking completed successfully');

    } catch (error) {
        console.error('标记基准文件失败:', error);
        alert('标记基准文件失败: ' + error.message);

        // 显示错误信息
        baseFileMarkedContent.innerHTML = '<p class="text-red-500">标记失败，请重试</p>';
    } finally {
        // 恢复按钮状态
        markBaseButton.disabled = false;
        hideLoading();
    }
}

// 标记对比文件修改
async function markCompareFileChanges() {
    console.log('markCompareFileChanges called:', {
        diffResult: !!diffResult,
        diffResultDiff: !!(diffResult && diffResult.diff),
        compareContent: !!(diffResult && diffResult.compareContent),
        compareFileMarkedContent: !!compareFileMarkedContent
    });

    if (!diffResult || !diffResult.diff || !diffResult.compareContent) return;

    // 显示加载状态
    markCompareButton.disabled = true;
    showLoading('正在标记对比文件...');

    try {
        // 重置标记状态
        compareFileMarked = false;

        // 创建标记内容 - 基于HTML内容进行标记
        let markedHtml = compareFileContent.html || '';

        // 简化版本：直接使用文本标记方法，但包装在HTML标签中
        let markedContent = '';
        diffResult.diff.forEach(part => {
            if (part.added) {
                markedContent += `<span class="added-text">${escapeHtml(part.value)}</span>`;
            } else if (!part.removed) {
                markedContent += escapeHtml(part.value);
            }
        });
        // 包装在段落标签中，确保DOCX生成能正确解析
        markedHtml = `<p>${markedContent.replace(/\n\n+/g, '</p><p>').replace(/\n/g, '<br>')}</p>`;

        // 显示标记内容
        console.log('Setting compare marked content:', markedHtml.substring(0, 500) + '...');
        console.log('Content contains modified-text:', markedHtml.includes('modified-text'));
        console.log('Content contains added-text:', markedHtml.includes('added-text'));
        compareFileMarkedContent.innerHTML = markedHtml;

        // 保存标记内容到compareFileContent对象
        compareFileContent.markedHtml = markedHtml;
        compareFileContent.markedText = markedHtml;

        // 更新状态
        compareFileMarked = true;
        compareFileStatus.textContent = '已标记';
        compareFileStatus.classList.add('text-accent', 'font-medium');

        // 滚动到结果区域
        compareFileMarkedContent.scrollIntoView({ behavior: 'smooth', block: 'start' });

        // 更新按钮状态
        updateButtonStates();

        // 保存内容以供下载
        saveContentForDownload('compare', '对比文件');

        console.log('Compare file marking completed successfully');

    } catch (error) {
        console.error('标记对比文件失败:', error);
        alert('标记对比文件失败: ' + error.message);

        // 显示错误信息
        compareFileMarkedContent.innerHTML = '<p class="text-red-500">标记失败，请重试</p>';
    } finally {
        // 恢复按钮状态
        markCompareButton.disabled = false;
        hideLoading();
    }
}

// HTML内容标记辅助函数 - 使用更保守的方法
function markHtmlContent(htmlContent, diffResult, markType) {
    try {
        // 如果HTML太复杂（包含表格），暂时回退到文本标记
        // 这是保守策略，确保标记功能正常工作
        if (htmlContent.includes('<table>')) {
            console.log('HTML contains table, using text-based marking to preserve functionality');
            // 对复杂HTML，使用纯文本差异方法，但保持基本格式
            return createTextBasedMarkedContent(htmlContent, diffResult, markType);
        }

        // 对于简单HTML，尝试直接标记
        return markHtmlWithTextDiff(htmlContent, diffResult, markType);

    } catch (error) {
        console.warn('HTML标记失败，使用文本标记:', error);
        return createTextBasedMarkedContent(htmlContent, diffResult, markType);
    }
}

// 获取节点的路径用于标记
function getNodePath(node) {
    const path = [];
    let current = node;
    while (current && current !== document.body) {
        if (!current.parentNode) {
            break; // 防止访问null的parentNode
        }
        const index = Array.from(current.parentNode.childNodes).indexOf(current);
        path.unshift({
            tagName: current.tagName,
            index: index
        });
        current = current.parentNode;
    }
    return path;
}

// 创建基于文本差异的标记内容（修复版本）
function createTextBasedMarkedContent(htmlContent, diff, markType) {
    console.log('Using text-based marking for complex HTML');

    // 过滤出相关的差异部分
    const relevantDiff = diff.filter(part =>
        (markType === 'removed' && part.removed) ||
        (markType === 'added' && part.added)
    );

    console.log(`Found ${relevantDiff.length} ${markType} changes`);

    // 如果变化部分太多，说明是逐字符级别的差异，返回包含完整结构的HTML
    if (relevantDiff.length > 50) {
        console.log('Too many changes detected (likely character-level), returning structured HTML content');
        // 对于太多变化的文档，保留原始HTML结构但确保有基本的段落结构
        if (htmlContent.includes('<p>') || htmlContent.includes('<div>')) {
            return htmlContent; // 返回原始HTML
        } else {
            // 如果没有段落结构，添加段落标签
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlContent;
            const plainText = tempDiv.textContent || tempDiv.innerText || htmlContent;
            return '<p>' + plainText.replace(/\n\n+/g, '</p><p>') + '</p>';
        }
    }

    // 如果没有找到任何变化，返回原始HTML内容
    if (relevantDiff.length === 0) {
        console.log(`No ${markType} changes found, returning original HTML content`);
        // 确保HTML有适当的结构
        if (htmlContent.includes('<p>') || htmlContent.includes('<div>') || htmlContent.includes('<table>')) {
            return htmlContent; // 返回原始HTML
        } else {
            // 为纯文本添加段落结构
            return '<p>' + htmlContent.replace(/\n\n+/g, '</p><p>') + '</p>';
        }
    }

    // 对于合理数量的变化，进行标记
    let resultContent = '';
    let foundChanges = 0;

    diff.forEach(part => {
        if ((markType === 'removed' && part.removed)) {
            // 只标记长度大于2的内容（避免标记单个字符）
            if (part.value.length > 2) {
                resultContent += `<span class="modified-text">${escapeHtml(part.value)}</span>`;
                foundChanges++;
            } else {
                resultContent += escapeHtml(part.value);
            }
        } else if ((markType === 'added' && part.added)) {
            // 只标记长度大于2的内容（避免标记单个字符）
            if (part.value.length > 2) {
                resultContent += `<span class="added-text">${escapeHtml(part.value)}</span>`;
                foundChanges++;
            } else {
                resultContent += escapeHtml(part.value);
            }
        } else if (!part.added && !part.removed) {
            // 保留未修改内容
            resultContent += escapeHtml(part.value);
        }
    });

    console.log(`Actually marked ${foundChanges} significant changes`);

    // 确保返回的内容有适当的HTML结构
    let finalContent = resultContent.replace(/\n\n+/g, '</p><p>').replace(/\n/g, '<br>');

    // 如果内容没有HTML标签，包装在段落中
    if (!finalContent.includes('<p>') && !finalContent.includes('<div>') && !finalContent.includes('<table>')) {
        finalContent = '<p>' + finalContent + '</p>';
    }

    return finalContent;
}

// 使用文本差异标记HTML内容（修复版本 - 合并小片段）
function markHtmlWithTextDiff(htmlContent, diff, markType) {
    // 对于HTML内容，我们使用更保守的方法
    console.log('HTML marking detected too many parts, using fallback approach');

    // 过滤出相关的差异部分
    const relevantDiff = diff.filter(part =>
        (markType === 'removed' && part.removed) ||
        (markType === 'added' && part.added)
    );

    console.log(`Found ${relevantDiff.length} ${markType} parts`);

    // 如果差异部分太多（逐字符级别），返回原始HTML内容
    if (relevantDiff.length > 100) {
        console.log('Too many small parts detected, returning original HTML');
        return htmlContent; // 直接返回原始HTML，保留表格结构
    }

    // 否则尝试在HTML中直接替换
    let markedContent = htmlContent;
    let successfulReplacements = 0;

    relevantDiff.forEach((part, index) => {
        const escapedValue = escapeHtml(part.value);

        // 跳过太短的片段（可能是标点符号或空格）
        if (escapedValue.length < 3) {
            return;
        }

        if ((markType === 'removed' && part.removed)) {
            const replacement = `<span class="modified-text">${escapedValue}</span>`;
            if (markedContent.includes(escapedValue)) {
                markedContent = markedContent.replace(escapedValue, replacement);
                successfulReplacements++;
            }
        } else if ((markType === 'added' && part.added)) {
            const replacement = `<span class="added-text">${escapedValue}</span>`;
            if (markedContent.includes(escapedValue)) {
                markedContent = markedContent.replace(escapedValue, replacement);
                successfulReplacements++;
            }
        }
    });

    console.log(`Successfully made ${successfulReplacements} replacements`);

    // 如果没有成功的替换，返回原始HTML内容
    if (successfulReplacements === 0) {
        console.log('No successful replacements, returning original HTML');
        return htmlContent;
    }

    return markedContent;
}

// 在文件标记完成时保存内容到全局变量供下载使用
function saveContentForDownload(type, fileName) {
    console.log('saveContentForDownload called:', { type, fileName });

    // 获取原始内容
    const originalContent = type === 'base' ? baseFileContent.text : compareFileContent.text;

    // 优先使用保存的标记内容，然后从DOM获取
    const fileContent = type === 'base' ? baseFileContent : compareFileContent;
    let actualMarkedContent = fileContent.markedHtml || fileContent.markedText || '';

    if (!actualMarkedContent) {
        // 从DOM元素获取标记内容（包含格式）
        const contentElement = type === 'base' ? baseFileMarkedContent : compareFileMarkedContent;

        if (contentElement && contentElement.innerHTML) {
            // 如果内容是"点击按钮查看标记内容"之类的占位符，不保存
            if (contentElement.innerHTML.includes('点击') || contentElement.innerHTML.includes('请') ||
                contentElement.innerHTML.includes('italic') || contentElement.innerHTML.includes('gray')) {
                console.log('检测到占位符内容，使用原始内容');
                actualMarkedContent = originalContent;
            } else {
                actualMarkedContent = contentElement.innerHTML;
                console.log('使用DOM标记内容，长度:', actualMarkedContent.length);
            }
        } else {
            actualMarkedContent = originalContent;
            console.log('DOM内容为空，使用原始内容');
        }
    } else {
        console.log('使用保存的标记内容，长度:', actualMarkedContent.length);
    }

    // 保存差异结果到全局变量供HTML中的DOCX生成函数使用
    if (typeof window !== 'undefined') {
        window.diffResult = diffResult;
        window.baseFileContent = baseFileContent;
        window.compareFileContent = compareFileContent;
    }

    // 调用HTML中定义的全局函数保存内容
    if (window.saveFileContents) {
        window.saveFileContents(type, originalContent, actualMarkedContent, fileName);
        console.log(`已保存${type}文件内容到全局变量`);
    } else {
        console.error('window.saveFileContents函数未找到');
    }
}


// HTML转义
function escapeHtml(text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

// 初始化应用
function initApp() {
    // 应用初始化时检查Diff库是否正确加载
    checkDiffLibrary();
    initEventListeners();
    updateButtonStates();
}

// 当DOM加载完成后初始化应用
document.addEventListener('DOMContentLoaded', initApp);