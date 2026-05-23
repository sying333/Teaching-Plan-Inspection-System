// 全局状态
let currentFile = null;
let currentFilePath = null;
let imageFolderPath = '';
let checkResult = null;
let highlightEnabled = true;

// 解析相关状态
let parsedMdContent = null;  // 解析后的md内容
let parsedMdPath = null;     // 解析后的md文件路径
let parsedImagesDir = null;  // 解析后的图片目录
let isParsedDocument = false; // 是否为解析后的文档

// DOM元素
const fileInput = document.getElementById('file-input');
const fileNameDisplay = document.getElementById('file-name');
const checkBtn = document.getElementById('check-btn');
const documentContent = document.getElementById('document-content');
const reportContent = document.getElementById('report-content');
const loadingOverlay = document.getElementById('loading-overlay');
const toggleHighlightBtn = document.getElementById('toggle-highlight');
const exportBtn = document.getElementById('export-btn');
const viewRulesBtn = document.getElementById('view-rules-btn');
const rulesModal = document.getElementById('rules-modal');
const closeRulesModal = document.getElementById('close-rules-modal');
const clearFolderBtn = document.getElementById('clear-folder-btn');
const imageFolderPathInput = document.getElementById('image-folder-path');

// 解析相关DOM元素
const parseStatus = document.getElementById('parse-status');
const parseProgressBar = document.getElementById('parse-progress-bar');
const parseMessage = document.getElementById('parse-message');
const parsedInfo = document.getElementById('parsed-info');
const imageFolderSelector = document.getElementById('image-folder-selector');

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    loadRecentFiles();
    loadRulesCount();
});

// 初始化事件监听
function initializeEventListeners() {
    fileInput.addEventListener('change', handleFileSelect);
    checkBtn.addEventListener('click', startCheck);
    toggleHighlightBtn.addEventListener('click', toggleHighlight);
    exportBtn.addEventListener('click', exportReport);
    viewRulesBtn.addEventListener('click', showRulesModal);
    closeRulesModal.addEventListener('click', hideRulesModal);
    clearFolderBtn.addEventListener('click', clearImageFolderPath);
    
    // 输入框改变时更新全局变量
    imageFolderPathInput.addEventListener('input', (e) => {
        imageFolderPath = e.target.value.trim();
    });
    
    // 点击模态框外部关闭
    rulesModal.addEventListener('click', (e) => {
        if (e.target === rulesModal) {
            hideRulesModal();
        }
    });

    // 添加高亮状态改变监听
    document.addEventListener('highlightsUpdated', () => {
        if (checkResult && checkResult.violations) {
            applyViolationHighlights(checkResult.violations);
        }
    });
}

// 文件选择处理
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    currentFile = file;
    currentFilePath = file.path || file.name;
    
    // 重置解析状态
    resetParseState();
    
    // 更新UI
    fileNameDisplay.textContent = file.name;
    
    // 更新文件信息
    showFileInfo(file);
    
    // 保存到最近使用
    saveToRecentFiles(file);
    
    // 根据文件类型决定处理方式
    const fileExt = file.name.toLowerCase().split('.').pop();
    
    if (fileExt === 'md') {
        // Markdown文件，直接读取显示
        isParsedDocument = false;
        imageFolderSelector.style.display = 'block';
        checkBtn.disabled = false;
        readAndDisplayFile(file);
    } else if (fileExt === 'docx' || fileExt === 'doc' || fileExt === 'pdf') {
        // Word/PDF文件，需要上传解析
        isParsedDocument = true;
        imageFolderSelector.style.display = 'none';  // 不需要手动指定图片目录
        checkBtn.disabled = true;  // 解析完成后才能检测
        uploadAndParseDocument(file);
    } else {
        alert('不支持的文件格式，请选择 .docx, .pdf 或 .md 文件');
        checkBtn.disabled = true;
    }
}

// 重置解析状态
function resetParseState() {
    parsedMdContent = null;
    parsedMdPath = null;
    parsedImagesDir = null;
    isParsedDocument = false;
    
    // 隐藏解析相关UI
    if (parseStatus) parseStatus.style.display = 'none';
    if (parsedInfo) parsedInfo.style.display = 'none';
    if (parseProgressBar) parseProgressBar.style.width = '0%';
}

// 上传并解析文档（docx/pdf）
async function uploadAndParseDocument(file) {
    console.log('开始上传解析文档:', file.name);
    
    // 显示解析状态
    parseStatus.style.display = 'block';
    parseMessage.textContent = '正在上传文档...';
    parseProgressBar.style.width = '10%';
    
    // 显示加载中
    documentContent.innerHTML = `
        <div class="empty-state">
            <div class="spinner"></div>
            <p>正在解析文档...</p>
            <p class="hint">docx/pdf文件需要通过MinerU模型解析，请稍候</p>
        </div>
    `;
    
    let parseSuccess = false;
    
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        parseMessage.textContent = '正在调用MinerU模型解析...';
        parseProgressBar.style.width = '30%';
        
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });
        
        console.log('收到响应, status:', response.status);
        
        // 尝试读取完整响应文本
        const responseText = await response.text();
        console.log('响应内容:', responseText.substring(0, 500));
        
        // 解析响应
        const lines = responseText.split('\n').filter(line => line.trim());
        
        for (const line of lines) {
            try {
                const data = JSON.parse(line);
                console.log('解析到数据:', data.type, data.success);
                
                if (data.type === 'complete' && data.success) {
                    // 调用新的处理函数
                    handleDocumentParsed(data);

                    parseSuccess = true;
                    
                    // 解析成功
                    parsedMdContent = data.md_content;
                    parsedMdPath = data.md_path;
                    parsedImagesDir = data.page_images_dir || data.images_dir;
                    
                    // 更新UI
                    parseStatus.style.display = 'none';
                    parsedInfo.style.display = 'block';
                    
                    document.getElementById('parsed-doc-name').textContent = data.doc_name;
                    document.getElementById('parsed-image-count').textContent = data.image_count + ' 页';
                    
                    // 显示VL分析信息
                    const vlInfo = data.vl_special_layouts > 0 
                        ? `${data.vl_special_layouts} 个特殊布局已分析`
                        : '无特殊布局';
                    document.getElementById('parsed-images-dir').textContent = vlInfo;
                    
                    // 自动设置图片目录（使用页面图片目录）
                    imageFolderPath = data.page_images_dir || data.images_dir;
                    imageFolderPathInput.value = imageFolderPath;
                    
                    // 启用检测按钮
                    checkBtn.disabled = false;
                    
                    console.log('双路径解析成功:', data.doc_name, '页数:', data.image_count, 'VL特殊布局:', data.vl_special_layouts);
                } else if (data.type === 'error') {
                    throw new Error(data.error || '解析失败');
                }
            } catch (parseError) {
                if (parseError.message && !parseError.message.includes('JSON')) {
                    throw parseError; // 重新抛出非JSON解析错误
                }
                console.error('JSON解析失败:', parseError, '原始行:', line.substring(0, 100));
            }
        }
        
        // 如果没有成功解析，显示错误
        if (!parseSuccess) {
            throw new Error('未收到有效的解析结果');
        }
        
    } catch (error) {
        console.error('上传解析失败:', error);
        parseStatus.style.display = 'none';
        documentContent.innerHTML = `
            <div class="empty-state error">
                <div class="empty-icon">失败</div>
                <p>文档解析失败</p>
                <p class="hint">${error.message}</p>
            </div>
        `;
        alert('文档解析失败: ' + error.message);
    }
}

// 显示解析后的内容
function displayParsedContent(mdContent) {
    if (!mdContent) {
        documentContent.innerHTML = '<div class="empty-state"><p>无法显示内容</p></div>';
        return;
    }
    
    // 处理内容
    let processedContent = mdContent;
    
    // 将<page N>标签转换为页面分隔符
    processedContent = processedContent.replace(/<page\s+(\d+)>/gi, (match, pageNum) => {
        return `\n\n---\n\n**📄 第 ${pageNum} 页**\n\n`;
    });
    
    // 移除闭合的</page N>标签
    processedContent = processedContent.replace(/<\/page\s*\d*>/gi, '\n\n');
    
    // 处理图片路径 - 对于解析后的文档，显示图片占位符
    processedContent = processedContent.replace(/!\[(.*?)\]\((.*?)\)/g, (match, alt, src) => {
        // 提取图片文件名
        const imgName = src.split('/').pop();
        return `\n\n**[图片内容: ${alt || imgName}]**\n\n`;
    });
    
    // 使用marked渲染Markdown
    const renderedHtml = marked.parse(processedContent);
    
    // 添加文档样式包装
    documentContent.innerHTML = `<div class="markdown-body rendered-document">${renderedHtml}</div>`;
    
    // 添加段落ID
    addParagraphIds();
}

// 清空图片文件夹路径
function clearImageFolderPath() {
    imageFolderPath = '';
    imageFolderPathInput.value = '';
    console.log('已清空图片文件夹路径');
}

// 显示文件信息
function showFileInfo(file) {
    const fileInfo = document.getElementById('file-info');
    const infoFilename = document.getElementById('info-filename');
    const infoFilesize = document.getElementById('info-filesize');
    const infoFiletype = document.getElementById('info-filetype');
    
    fileInfo.style.display = 'block';
    infoFilename.textContent = file.name;
    infoFilesize.textContent = formatFileSize(file.size);
    
    // 显示文件类型
    const fileExt = file.name.toLowerCase().split('.').pop();
    const typeMap = {
        'md': 'Markdown文档',
        'docx': 'Word文档',
        'doc': 'Word文档',
        'pdf': 'PDF文档'
    };
    infoFiletype.textContent = typeMap[fileExt] || '未知类型';
}

// 格式化文件大小
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// 读取并显示文件
function readAndDisplayFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const markdown = e.target.result;
        renderMarkdown(markdown);
    };
    reader.readAsText(file);
}

// 渲染Markdown
function renderMarkdown(markdown, violations = []) {
    // 处理内容
    let processedContent = markdown;
    
    // 将<page N>标签转换为页面分隔符
    processedContent = processedContent.replace(/<page\s+(\d+)>/gi, (match, pageNum) => {
        return `\n\n---\n\n**📄 第 ${pageNum} 页**\n\n`;
    });
    
    // 移除闭合的</page N>标签
    processedContent = processedContent.replace(/<\/page\s*\d*>/gi, '\n\n');
    
    // 处理图片语法，替换为图片占位符
    processedContent = processedContent.replace(/!\[(.*?)\]\((.*?)\)/g, (match, alt, src) => {
        const imgName = src.split('/').pop();
        return `\n\n**[图片内容: ${alt || imgName}]**\n\n`;
    });
    
    // 使用marked渲染
    const html = marked.parse(processedContent);
    
    // 添加文档样式包装
    documentContent.innerHTML = `<div class="markdown-body rendered-document">${html}</div>`;
    
    // 等待DOM渲染完成后再添加段落ID和高亮
    setTimeout(() => {
        // 添加段落ID
        addParagraphIds();
        
        // 如果有违规信息，应用高亮
        if (violations.length > 0) {
            console.log('渲染后应用违规高亮，违规数:', violations.length);
            applyViolationHighlights(violations);
        }
    }, 100);
}

// 添加段落ID
function addParagraphIds() {
    const elements = documentContent.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td');
    elements.forEach((el, index) => {
        el.setAttribute('data-paragraph-id', `para-${index}`);
    });
}

// 应用违规高亮
function applyViolationHighlights(violations) {
    console.log('应用违规高亮开始，违规数:', violations.length, '高亮开关:', highlightEnabled);
    
    if (!highlightEnabled) {
        console.log('高亮已关闭，跳过应用高亮');
        return;
    }
    
    // 清除之前的高亮
    const oldHighlights = document.querySelectorAll('.violation-highlight');
    oldHighlights.forEach(el => el.remove());
    
    // 重新应用高亮
    violations.forEach((violation, index) => {
        setTimeout(() => {
            applyViolationHighlight(violation, index);
        }, index * 10); // 稍微错开时间，避免阻塞
    });
}

// 应用单个违规高亮
function applyViolationHighlight(violation, index) {
    // 获取违规内容
    let content = violation.content || violation.segment_data?.content || '';
    content = stripHtmlTags(content).trim();
    
    if (!content || content.length < 3) {
        console.log(`违规${index}: 内容太短或为空，跳过`);
        return;
    }
    
    console.log(`违规${index}: 搜索内容="${content.substring(0, 50)}..."`);
    
    // 在文档中查找匹配的文本
    const documentContainer = document.querySelector('.rendered-document');
    if (!documentContainer) {
        console.log('未找到文档容器');
        return;
    }
    
    // 使用TreeWalker查找文本节点
    const walker = document.createTreeWalker(
        documentContainer,
        NodeFilter.SHOW_TEXT,
        null,
        false
    );
    
    const textNodes = [];
    let node;
    while (node = walker.nextNode()) {
        textNodes.push(node);
    }
    
    // 查找包含违规内容的文本节点
    for (const textNode of textNodes) {
        const nodeText = textNode.textContent;
        const searchIndex = nodeText.indexOf(content);
        
        if (searchIndex !== -1) {
            // 找到匹配，高亮该文本节点
            highlightTextNode(textNode, searchIndex, content.length, violation, index);
            return;
        }
        
        // 尝试模糊匹配（忽略多余空格）
        const nodeTextCompact = nodeText.replace(/\s+/g, '');
        const contentCompact = content.replace(/\s+/g, '');
        
        if (nodeTextCompact.includes(contentCompact) && contentCompact.length > 5) {
            // 找到模糊匹配
            highlightTextNodeFuzzy(textNode, content, violation, index);
            return;
        }
    }
    
    console.log(`违规${index}: 未找到匹配的文本节点`);
}

// 高亮文本节点
function highlightTextNode(textNode, startIndex, length, violation, index) {
    const parent = textNode.parentNode;
    const text = textNode.textContent;
    
    const beforeText = text.substring(0, startIndex);
    const highlightText = text.substring(startIndex, startIndex + length);
    const afterText = text.substring(startIndex + length);
    
    // 创建高亮元素
    const highlightSpan = document.createElement('span');
    highlightSpan.className = 'violation-highlight';
    highlightSpan.setAttribute('data-violation-index', index);
    highlightSpan.textContent = highlightText;
    
    // 添加高亮样式类
    highlightSpan.classList.add('highlight-visible');
    
    // 添加点击事件
    highlightSpan.addEventListener('click', (e) => {
        e.stopPropagation();
        expandViolationItem(index);
    });
    
    // 添加鼠标悬停事件
    highlightSpan.addEventListener('mouseenter', (e) => {
        showTooltip(e, violation);
    });
    
    highlightSpan.addEventListener('mouseleave', () => {
        hideTooltip();
    });
    
    // 替换文本节点
    const fragment = document.createDocumentFragment();
    if (beforeText) fragment.appendChild(document.createTextNode(beforeText));
    fragment.appendChild(highlightSpan);
    if (afterText) fragment.appendChild(document.createTextNode(afterText));
    
    parent.replaceChild(fragment, textNode);
    
    console.log(`违规${index}: 成功高亮文本"${highlightText.substring(0, 30)}..."`);
}

// 模糊高亮文本节点
function highlightTextNodeFuzzy(textNode, searchText, violation, index) {
    // 简化实现：高亮整个文本节点的父元素
    const parentElement = textNode.parentElement;
    
    if (parentElement && parentElement.tagName !== 'BODY') {
        parentElement.classList.add('violation-highlight');
        parentElement.setAttribute('data-violation-index', index);
        
        // 添加高亮样式类
        parentElement.classList.add('highlight-visible');
        
        // 添加事件
        parentElement.addEventListener('click', (e) => {
            e.stopPropagation();
            expandViolationItem(index);
        });
        
        console.log(`违规${index}: 模糊高亮整个元素`);
    }
}

// 高亮文本 - 精确到句子级别
function highlightText(searchText, violation, index) {
    // 清理搜索文本
    const cleanSearch = searchText.trim();
    if (cleanSearch.length < 5) return;
    
    console.log(`违规${index}: 搜索文本="${cleanSearch.substring(0, 30)}..."`);
    
    // 获取所有可能包含文本的元素
    const allElements = documentContent.querySelectorAll('p, td, th, li, h1, h2, h3, h4, h5, h6, blockquote');
    
    let foundElement = null;
    let matchedText = null;
    
    // 遍历所有元素，查找包含违规文本的元素
    for (const el of allElements) {
        const elText = el.textContent;
        
        // 尝试多种匹配方式
        // 1. 完全匹配
        if (elText.includes(cleanSearch)) {
            foundElement = el;
            matchedText = cleanSearch;
            console.log(`违规${index}: 完全匹配成功`);
            break;
        }
        
        // 2. 忽略空格匹配
        const elTextNoSpace = elText.replace(/\s+/g, '');
        const searchNoSpace = cleanSearch.replace(/\s+/g, '');
        if (elTextNoSpace.includes(searchNoSpace)) {
            foundElement = el;
            // 尝试恢复原始匹配文本
            matchedText = findOriginalText(elText, searchNoSpace);
            console.log(`违规${index}: 忽略空格匹配成功`);
            break;
        }
    }
    
    // 如果找到匹配的元素，在元素内精确高亮句子
    if (foundElement && matchedText) {
        highlightExactText(foundElement, matchedText, violation, index);
    } else {
        console.log(`违规${index}: 未找到匹配`);
    }
}

// 在原文中找到匹配的文本（保留原始空格）
function findOriginalText(originalText, searchNoSpace) {
    // 尝试找到原文中对应的部分
    let result = '';
    let searchIdx = 0;
    
    for (let i = 0; i < originalText.length && searchIdx < searchNoSpace.length; i++) {
        const char = originalText[i];
        if (/\s/.test(char)) {
            result += char;
        } else if (char === searchNoSpace[searchIdx]) {
            result += char;
            searchIdx++;
        } else {
            // 不匹配，重新开始
            result = '';
            searchIdx = 0;
            if (char === searchNoSpace[0]) {
                result = char;
                searchIdx = 1;
            }
        }
    }
    
    return searchIdx === searchNoSpace.length ? result.trim() : searchNoSpace;
}

// 在元素内精确高亮指定文本
function highlightExactText(element, textToHighlight, violation, index) {
    const ruleName = violation.violated_rules?.[0]?.rule_name || '违规内容';
    
    // 使用TreeWalker遍历文本节点
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
    const textNodes = [];
    let node;
    
    while (node = walker.nextNode()) {
        textNodes.push(node);
    }
    
    // 合并所有文本内容
    const fullText = textNodes.map(n => n.textContent).join('');
    
    // 查找匹配位置
    const matchStart = fullText.indexOf(textToHighlight);
    if (matchStart === -1) {
        // 尝试忽略空格匹配
        const fullTextNoSpace = fullText.replace(/\s+/g, '');
        const searchNoSpace = textToHighlight.replace(/\s+/g, '');
        if (fullTextNoSpace.includes(searchNoSpace)) {
            // 整个元素高亮作为备用
            applyHighlightToElement(element, violation, index);
        }
        return;
    }
    
    const matchEnd = matchStart + textToHighlight.length;
    
    // 找到匹配文本所在的文本节点并包裹
    let currentPos = 0;
    let highlightApplied = false;
    
    for (const textNode of textNodes) {
        const nodeStart = currentPos;
        const nodeEnd = currentPos + textNode.textContent.length;
        
        // 检查这个文本节点是否包含匹配的文本
        if (nodeEnd > matchStart && nodeStart < matchEnd) {
            // 计算在这个节点内的起止位置
            const localStart = Math.max(0, matchStart - nodeStart);
            const localEnd = Math.min(textNode.textContent.length, matchEnd - nodeStart);
            
            // 分割文本节点
            const before = textNode.textContent.substring(0, localStart);
            const match = textNode.textContent.substring(localStart, localEnd);
            const after = textNode.textContent.substring(localEnd);
            
            // 创建高亮span
            const span = document.createElement('span');
            span.className = 'violation-highlight';
            span.setAttribute('data-violation-index', index);
            span.setAttribute('data-violation-type', ruleName);
            span.textContent = match;
            
            // 添加事件
            span.addEventListener('click', (e) => {
                e.stopPropagation();
                expandViolationItem(index);
            });
            span.addEventListener('mouseenter', (e) => showTooltip(e, violation));
            span.addEventListener('mouseleave', () => hideTooltip());
            
            // 替换原文本节点
            const fragment = document.createDocumentFragment();
            if (before) fragment.appendChild(document.createTextNode(before));
            fragment.appendChild(span);
            if (after) fragment.appendChild(document.createTextNode(after));
            
            textNode.parentNode.replaceChild(fragment, textNode);
            highlightApplied = true;
            
            console.log(`违规${index}: 精确高亮应用成功，文本="${match.substring(0, 20)}..."`);
            break; // 只高亮第一个匹配
        }
        
        currentPos = nodeEnd;
    }
    
    if (!highlightApplied) {
        // 备用：高亮整个元素
        applyHighlightToElement(element, violation, index);
    }
}

// 应用高亮到整个元素（备用方法）
function applyHighlightToElement(el, violation, index) {
    el.classList.add('violation-highlight');
    el.setAttribute('data-violation-index', index);
    el.setAttribute('data-violation-type', violation.violated_rules?.[0]?.rule_name || '违规内容');
    
    // 添加点击事件
    el.addEventListener('click', (e) => {
        e.stopPropagation();
        expandViolationItem(index);
    });
    
    // 添加悬浮事件
    el.addEventListener('mouseenter', (e) => showTooltip(e, violation));
    el.addEventListener('mouseleave', () => hideTooltip());
}

// 显示提示
function showTooltip(event, violation) {
    const tooltip = document.createElement('div');
    tooltip.className = 'violation-tooltip show';
    tooltip.textContent = violation.violated_rules[0]?.rule_name || '违规内容';
    tooltip.style.left = event.pageX + 'px';
    tooltip.style.top = (event.pageY - 30) + 'px';
    document.body.appendChild(tooltip);
    
    setTimeout(() => tooltip.classList.add('show'), 10);
}

// 隐藏提示
function hideTooltip() {
    const tooltips = document.querySelectorAll('.violation-tooltip');
    tooltips.forEach(t => t.remove());
}

// 开始检测
async function startCheck() {
    if (!currentFile && !parsedMdContent) return;
    
    // 显示加载动画
    loadingOverlay.style.display = 'flex';
    checkBtn.disabled = true;
    
    // 重置进度条
    updateProgress(0, '正在准备检测...');
    
    try {
        let fileContent;
        let filename;
        let imageFolder;
        
        if (isParsedDocument && parsedMdContent) {
            // 使用解析后的内容
            updateProgress(10, '使用解析后的内容...');
            fileContent = parsedMdContent;
            filename = currentFile ? currentFile.name.replace(/\.(docx|doc|pdf)$/i, '.md') : 'document.md';
            imageFolder = parsedImagesDir || imageFolderPath || 'images';
            console.log('使用解析后的文档，图片目录:', imageFolder);
        } else {
            // 读取原始md文件内容
            updateProgress(10, '正在读取文件内容...');
            fileContent = await readFileContent(currentFile);
            filename = currentFile.name;
            imageFolder = imageFolderPath || 'images';
        }
        
        // 调试信息
        console.log('文件长度:', fileContent.length);
        console.log('图片目录:', imageFolder);
        
        updateProgress(15, '正在连接服务器...');
        
        // 调用API（流式响应）
        const response = await fetch('/api/check', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                content: fileContent,
                filename: filename,
                image_folder: imageFolder
            })
        });
        
        if (!response.ok) {
            let errorMsg = '检测请求失败';
            try {
                const err = await response.json();
                errorMsg = err.error || errorMsg;
            } catch(e) {}
            throw new Error(errorMsg);
        }
        
        // 处理流式响应
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            
            // 处理完整的行
            buffer = lines.pop(); // 保留最后一个可能不完整的行
            
            for (const line of lines) {
                if (!line.trim()) continue;
                
                try {
                    const data = JSON.parse(line);
                    
                    if (data.type === 'progress') {
                        updateProgress(data.percent, data.message);
                    } else if (data.type === 'result') {
                        const result = data.data;
                        checkResult = result;
                        
                        updateProgress(100, '检测完成！');
                        await sleep(500);
                        
                        // 渲染结果
                        renderMarkdown(result.original_content, result.violations);
                        renderReport(result);
                        
                        // 启用导出按钮
                        exportBtn.disabled = false;
                    } else if (data.type === 'error') {
                        throw new Error(data.message);
                    }
                } catch (e) {
                    console.error('解析响应数据失败:', e, line);
                }
            }
        }
        
    } catch (error) {
        console.error('检测错误:', error);
        alert('检测失败：' + error.message);
        updateProgress(0, '检测失败');
    } finally {
        loadingOverlay.style.display = 'none';
        checkBtn.disabled = false;
    }
}

// 更新进度条
function updateProgress(percent, status) {
    const progressBar = document.getElementById('progress-bar');
    const progressText = document.getElementById('progress-text');
    const loadingStatus = document.getElementById('loading-status');
    
    if (progressBar) {
        progressBar.style.width = percent + '%';
    }
    
    if (progressText) {
        progressText.textContent = percent + '%';
    }
    
    if (loadingStatus && status) {
        loadingStatus.textContent = status;
    }
}

// 延迟函数
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// 读取文件内容
function readFileContent(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(new Error('文件读取失败'));
        reader.readAsText(file);
    });
}

// 渲染报告
function renderReport(result) {
    const report = result.report;
    const violations = result.violations;
    
    // 确保统计数据存在
    if (!report.risk_level_count) {
        report.risk_level_count = { high: 0, medium: 0, low: 0 };
        violations.forEach(v => {
            const s = (v.severity || 'low').toLowerCase();
            if (report.risk_level_count[s] !== undefined) report.risk_level_count[s]++;
        });
    }
    
    // 计算总分和总体风险（如果后端未返回）
    if (report.overall_score === undefined) {
        const deductions = { high: 10, medium: 5, low: 2 };
        let deduction = 
            report.risk_level_count.high * deductions.high +
            report.risk_level_count.medium * deductions.medium +
            report.risk_level_count.low * deductions.low;
        report.overall_score = Math.max(0, 100 - deduction);
    }
    
    if (!report.risk_level) {
        if (report.risk_level_count.high > 0) report.risk_level = 'high';
        else if (report.risk_level_count.medium > 0) report.risk_level = 'medium';
        else report.risk_level = 'low';
    }

    // 总体评分
    const riskLevel = report.risk_level || 'low';
    const summaryHtml = `
        <div class="report-summary">
            <div class="score-display">
                <div class="score-number">${report.overall_score}</div>
                <div class="risk-level risk-${riskLevel.toLowerCase()}">${getRiskLevelText(riskLevel)}</div>
            </div>
            <div class="stats-grid">
                <div class="stat-item">
                    <span class="stat-value">${report.total_violations || violations.length}</span>
                    <span class="stat-label">违规项</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${report.risk_level_count.high}</span>
                    <span class="stat-label">高风险</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${report.risk_level_count.medium}</span>
                    <span class="stat-label">中风险</span>
                </div>
                <div class="stat-item">
                    <span class="stat-value">${report.risk_level_count.low}</span>
                    <span class="stat-label">低风险</span>
                </div>
            </div>
        </div>
    `;
    
    // 违规列表
    let violationsHtml = '<div class="violations-list"><h3>违规详情</h3>';
    
    if (violations.length === 0) {
        violationsHtml += '<div class="no-violations">当前文档未发现明显违规项。</div>';
    } else {
        violations.forEach((violation, index) => {
            const rules = violation.violated_rules.map(r => 
                `<span class="rule-tag">${r.rule_name}</span>`
            ).join('');
            
            // 清理内容中的HTML标签
            const cleanContent = stripHtmlTags(violation.content || '');
            const cleanSuggestions = stripHtmlTags(violation.suggestions || '');
            const cleanLegalBasis = stripHtmlTags(violation.legal_basis || '');
            
            violationsHtml += `
                <div class="violation-item severity-${violation.severity}" id="violation-item-${index}">
                    <div class="violation-header" onclick="handleViolationClick(${index})">
                        <div>
                            <div class="violation-title">${violation.violated_rules[0]?.rule_name || '违规内容'}</div>
                            <div class="violation-location">第${violation.page_number}页 · 行${violation.line_range}</div>
                        </div>
                        <span class="violation-badge ${violation.severity}">${getSeverityText(violation.severity)}</span>
                    </div>
                    <div class="violation-body" id="violation-body-${index}">
                        <div class="violation-content">
                            ${escapeHtml(cleanContent)}
                        </div>
                        <div class="violation-rules">
                            <h4>违反规则:</h4>
                            ${rules}
                        </div>
                        ${cleanSuggestions ? `
                            <div class="violation-suggestion">
                                <h4>修改建议:</h4>
                                <p>${escapeHtml(cleanSuggestions)}</p>
                            </div>
                        ` : ''}
                        ${cleanLegalBasis ? `
                            <div class="violation-legal">
                                <h4>法律依据:</h4>
                                <p>${escapeHtml(cleanLegalBasis)}</p>
                            </div>
                        ` : ''}
                        ${violation.image_path ? `
                            <div class="violation-image-path">
                                <h4>图片路径:</h4>
                                <p>${escapeHtml(violation.image_path)}</p>
                            </div>
                        ` : ''}
                    </div>
                </div>
            `;
        });
    }
    
    violationsHtml += '</div>';
    
    reportContent.innerHTML = summaryHtml + violationsHtml;
}

// 切换违规项展开
function toggleViolation(index) {
    const body = document.getElementById(`violation-body-${index}`);
    if (body) {
        body.classList.toggle('expanded');
    }
}

// 滚动到违规项
function scrollToViolation(index) {
    const item = document.getElementById(`violation-item-${index}`);
    if (item) {
        item.scrollIntoView({ behavior: 'smooth', block: 'center' });
        
        // 展开
        const body = document.getElementById(`violation-body-${index}`);
        if (body) {
            body.classList.add('expanded');
        }
        
        // 闪烁提示
        item.style.animation = 'none';
        setTimeout(() => {
            item.style.animation = 'flash 0.5s ease-in-out 2';
        }, 10);
    }
}

// 闪烁元素
function flashElement(element) {
    element.classList.add('active');
    setTimeout(() => {
        element.classList.remove('active');
    }, 1500);
}

// 切换高亮
function toggleHighlight() {
    highlightEnabled = !highlightEnabled;
    const highlights = document.querySelectorAll('.violation-highlight');
    
    if (highlightEnabled) {
        toggleHighlightBtn.textContent = '关闭高亮';
        toggleHighlightBtn.classList.add('active');
        
        // 显示高亮
        highlights.forEach(el => {
            el.classList.add('highlight-visible');
            el.classList.remove('highlight-hidden');
        });
        
        console.log('高亮已开启');
    } else {
        toggleHighlightBtn.textContent = '开启高亮';
        toggleHighlightBtn.classList.remove('active');
        
        // 隐藏高亮
        highlights.forEach(el => {
            el.classList.add('highlight-hidden');
            el.classList.remove('highlight-visible');
        });
        
        console.log('高亮已关闭');
    }
}

// 导出报告
function exportReport() {
    if (!checkResult) return;
    
    const report = checkResult.report;
    const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${currentFile.name}_report.json`;
    a.click();
    URL.revokeObjectURL(url);
}

// 显示规则弹窗
async function showRulesModal() {
    rulesModal.style.display = 'flex';
    
    try {
        const response = await fetch('/api/rules');
        const data = await response.json();
        
        if (data.success) {
            renderRules(data.rules);
        }
    } catch (error) {
        console.error('加载规则失败:', error);
    }
}

// 隐藏规则弹窗
function hideRulesModal() {
    rulesModal.style.display = 'none';
}

// 渲染规则
function renderRules(rules) {
    const modalBody = document.getElementById('rules-modal-body');
    
    let html = '';
    rules.forEach(rule => {
        html += `
            <div class="rule-item">
                <h4>${rule.id}. ${rule.name}</h4>
                <p>${rule.description}</p>
                <div class="rule-meta">
                    <span class="rule-severity ${rule.severity}">${getSeverityText(rule.severity)}</span>
                    <span class="rule-category">${rule.category}</span>
                </div>
            </div>
        `;
    });
    
    modalBody.innerHTML = html;
}

// 加载规则数量
async function loadRulesCount() {
    try {
        const response = await fetch('/api/rules');
        const data = await response.json();
        
        if (data.success) {
            document.getElementById('rules-count').textContent = data.rules.length;
        }
    } catch (error) {
        console.error('加载规则数量失败:', error);
    }
}

// 最近文件相关
function loadRecentFiles() {
    const recent = JSON.parse(localStorage.getItem('recentFiles') || '[]');
    const list = document.getElementById('recent-files-list');
    
    if (recent.length === 0) {
        list.innerHTML = '<li class="empty-state">暂无历史记录</li>';
        return;
    }
    
    list.innerHTML = recent.map(file => `
        <li onclick="loadRecentFile('${file.path}')">${file.name}</li>
    `).join('');
}

function saveToRecentFiles(file) {
    let recent = JSON.parse(localStorage.getItem('recentFiles') || '[]');
    
    // 移除重复项
    recent = recent.filter(f => f.path !== (file.path || file.name));
    
    // 添加到开头
    recent.unshift({
        name: file.name,
        path: file.path || file.name,
        timestamp: Date.now()
    });
    
    // 只保留最近10个
    recent = recent.slice(0, 10);
    
    localStorage.setItem('recentFiles', JSON.stringify(recent));
    loadRecentFiles();
}

function loadRecentFile(path) {
    alert('请手动选择文件：' + path);
}

// 工具函数
function getRiskLevelText(level) {
    const map = {
        'low': '低风险',
        'medium': '中风险',
        'high': '高风险'
    };
    return map[level.toLowerCase()] || level;
}

function getSeverityText(severity) {
    const map = {
        'high': '高',
        'medium': '中',
        'low': '低'
    };
    return map[severity] || severity;
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// 清理HTML标签，只保留纯文本
function stripHtmlTags(text) {
    if (!text) return '';
    
    let cleaned = text;
    
    // 先解码HTML实体（如 &lt; &gt; 等）
    const textarea = document.createElement('textarea');
    textarea.innerHTML = cleaned;
    cleaned = textarea.value;
    
    // 移除所有HTML标签
    cleaned = cleaned.replace(/<[^>]*>/g, ' ');
    
    // 再次解码可能残留的HTML实体
    textarea.innerHTML = cleaned;
    cleaned = textarea.value;
    
    // 移除Markdown表格语法
    cleaned = cleaned.replace(/\|/g, ' ');
    
    // 清理多余空格和换行
    cleaned = cleaned.replace(/\s+/g, ' ').trim();
    
    return cleaned;
}

// 从中间栏滚动到文档中的违规位置
function scrollToDocumentViolation(violationIndex) {
    const element = documentContent.querySelector(`[data-violation-index="${violationIndex}"]`);
    if (element) {
        element.scrollIntoView({ behavior: 'smooth', block: 'center' });
        // 闪烁提示
        element.classList.add('flashing');
        setTimeout(() => element.classList.remove('flashing'), 2000);
    }
}

// 从右侧栏展开对应违规项
function expandViolationItem(violationIndex) {
    const item = document.getElementById(`violation-item-${violationIndex}`);
    const body = document.getElementById(`violation-body-${violationIndex}`);
    
    if (item) {
        item.scrollIntoView({ behavior: 'smooth', block: 'center' });
        // 闪烁提示
        item.classList.add('highlight-flash');
        setTimeout(() => item.classList.remove('highlight-flash'), 2000);
    }
    
    if (body) {
        body.classList.add('expanded');
    }
}

// 点击右侧栏违规项 → 展开详情 + 中间栏滚动到对应位置
function handleViolationClick(index) {
    // 展开详情
    toggleViolation(index);
    // 滚动到中间栏对应位置
    scrollToDocumentViolation(index);
}

// 使函数全局可访问
window.toggleViolation = toggleViolation;
window.handleViolationClick = handleViolationClick;
window.scrollToDocumentViolation = scrollToDocumentViolation;
window.expandViolationItem = expandViolationItem;

// 分页相关全局变量
let currentDocumentPages = [];
let currentPageIndex = 0;
let totalPages = 0;

function getPaginationElements() {
    return {
        paginationControls: document.getElementById('pagination-controls'),
        pageContentContainer: document.getElementById('page-content-container'),
        prevPageBtn: document.getElementById('prev-page-btn'),
        nextPageBtn: document.getElementById('next-page-btn'),
        currentPageLabel: document.getElementById('current-page'),
        totalPagesLabel: document.getElementById('total-pages'),
        pageInput: document.getElementById('page-input'),
        goToPageBtn: document.getElementById('go-to-page-btn')
    };
}

function getCurrentRenderedDocument() {
    return document.querySelector('#page-content-container .rendered-document')
        || document.querySelector('#document-content .rendered-document');
}

function preprocessMarkdownForPreview(markdownContent) {
    if (!markdownContent) {
        return '';
    }

    return markdownContent
        .replace(/!\[(.*?)\]\((.*?)\)/g, (match, alt, src) => {
            const imgName = src.split('/').pop();
            return `\n\n**[图片内容: ${alt || imgName}]**\n\n`;
        })
        .trim();
}

function resetDocumentPreview() {
    documentContent.innerHTML = `
        <div class="empty-state">
            <div class="empty-icon">文档</div>
            <p>请从左侧选择文档</p>
            <p class="hint">支持 Word (.docx) / PDF / Markdown (.md) 格式</p>
        </div>
    `;
}

// 重置解析状态，同时重置文档预览
function resetParseState() {
    parsedMdContent = null;
    parsedMdPath = null;
    parsedImagesDir = null;
    isParsedDocument = false;
    checkResult = null;

    if (parseStatus) parseStatus.style.display = 'none';
    if (parsedInfo) parsedInfo.style.display = 'none';
    if (parseProgressBar) parseProgressBar.style.width = '0%';
    if (exportBtn) exportBtn.disabled = true;

    initPagination();
    resetDocumentPreview();
}

// 分页功能初始化
function initPagination() {
    const { paginationControls, pageContentContainer, pageInput } = getPaginationElements();

    currentDocumentPages = [];
    currentPageIndex = 0;
    totalPages = 0;

    paginationControls.style.display = 'none';
    pageContentContainer.style.display = 'none';
    pageContentContainer.innerHTML = '';
    documentContent.style.display = 'block';

    if (pageInput) {
        pageInput.value = 1;
    }
}

// 解析Markdown内容并分割为页面
function parsePagesFromMarkdown(markdownContent) {
    const pages = [];
    const rawContent = markdownContent || '';
    const pageRegex = /<page\s+(\d+)>([\s\S]*?)<\/page\s*\d+>/gi;
    let match;

    while ((match = pageRegex.exec(rawContent)) !== null) {
        const pageNumber = parseInt(match[1], 10);
        const previewContent = preprocessMarkdownForPreview(match[2]);

        pages.push({
            pageNumber,
            content: previewContent,
            html: marked.parse(previewContent)
        });
    }

    if (pages.length === 0) {
        const previewContent = preprocessMarkdownForPreview(rawContent);
        pages.push({
            pageNumber: 1,
            content: previewContent,
            html: marked.parse(previewContent)
        });
    }

    return pages;
}

function renderCodeHighlight() {
    if (typeof hljs === 'undefined') {
        return;
    }

    document.querySelectorAll('#page-content-container pre code, #document-content pre code').forEach((block) => {
        hljs.highlightBlock(block);
    });
}

function addParagraphIds() {
    const renderedDocument = getCurrentRenderedDocument();
    if (!renderedDocument) {
        return;
    }

    const elements = renderedDocument.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, td, blockquote');
    elements.forEach((el, index) => {
        el.setAttribute('data-paragraph-id', `para-${index}`);
    });
}

function updatePaginationControls() {
    const {
        prevPageBtn,
        nextPageBtn,
        currentPageLabel,
        totalPagesLabel,
        pageInput
    } = getPaginationElements();

    const currentPage = currentDocumentPages[currentPageIndex];
    if (!currentPage) {
        return;
    }

    currentPageLabel.textContent = currentPage.pageNumber;
    totalPagesLabel.textContent = currentDocumentPages.length;
    pageInput.value = currentPage.pageNumber;
    pageInput.min = 1;
    pageInput.max = currentDocumentPages.length;
    prevPageBtn.disabled = currentPageIndex === 0;
    nextPageBtn.disabled = currentPageIndex === currentDocumentPages.length - 1;
}

// 显示特定页面
function showPage(pageIndex) {
    const { pageContentContainer } = getPaginationElements();

    if (pageIndex < 0 || pageIndex >= currentDocumentPages.length) {
        return;
    }

    currentPageIndex = pageIndex;
    const page = currentDocumentPages[pageIndex];

    updatePaginationControls();

    pageContentContainer.innerHTML = `
        <div class="page-section">
            <div class="page-header">
                <div class="page-title">第 ${page.pageNumber} 页</div>
                <div class="page-nav-hint">通过翻页按钮或键盘左右方向键切换页面</div>
            </div>
            <div class="page-content markdown-body rendered-document">
                ${page.html}
            </div>
        </div>
    `;

    addParagraphIds();
    renderCodeHighlight();

    if (checkResult?.violations?.length) {
        setTimeout(() => {
            applyViolationHighlights(checkResult.violations);
        }, 60);
    }

    pageContentContainer.scrollTop = 0;
}

function getPageIndexByNumber(pageNumber) {
    return currentDocumentPages.findIndex((page) => page.pageNumber === pageNumber);
}

// 初始化分页事件监听
function setupPaginationEvents() {
    const {
        prevPageBtn,
        nextPageBtn,
        pageInput,
        goToPageBtn
    } = getPaginationElements();

    if (prevPageBtn.dataset.bound === 'true') {
        return;
    }

    prevPageBtn.dataset.bound = 'true';

    prevPageBtn.addEventListener('click', () => {
        if (currentPageIndex > 0) {
            showPage(currentPageIndex - 1);
        }
    });

    nextPageBtn.addEventListener('click', () => {
        if (currentPageIndex < currentDocumentPages.length - 1) {
            showPage(currentPageIndex + 1);
        }
    });

    const goToPage = () => {
        if (!currentDocumentPages.length) {
            return;
        }

        let pageNumber = parseInt(pageInput.value, 10);
        if (Number.isNaN(pageNumber)) {
            pageNumber = currentDocumentPages[0].pageNumber;
        }

        const targetIndex = getPageIndexByNumber(pageNumber);
        if (targetIndex !== -1) {
            showPage(targetIndex);
            return;
        }

        const boundedIndex = Math.min(
            Math.max(pageNumber - 1, 0),
            currentDocumentPages.length - 1
        );
        showPage(boundedIndex);
    };

    goToPageBtn.addEventListener('click', goToPage);
    pageInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            goToPage();
        }
    });

    document.addEventListener('keydown', (e) => {
        if (!currentDocumentPages.length || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            return;
        }

        if (e.key === 'ArrowLeft') {
            e.preventDefault();
            if (currentPageIndex > 0) {
                showPage(currentPageIndex - 1);
            }
        }

        if (e.key === 'ArrowRight') {
            e.preventDefault();
            if (currentPageIndex < currentDocumentPages.length - 1) {
                showPage(currentPageIndex + 1);
            }
        }
    });
}

// 统一的文档预览更新入口
function updateDocumentContent(markdownContent) {
    const { paginationControls, pageContentContainer } = getPaginationElements();

    currentDocumentPages = parsePagesFromMarkdown(markdownContent);
    totalPages = currentDocumentPages.length;

    if (!totalPages) {
        initPagination();
        resetDocumentPreview();
        return;
    }

    documentContent.innerHTML = '';
    documentContent.style.display = 'none';
    pageContentContainer.style.display = 'block';
    paginationControls.style.display = totalPages > 1 ? 'flex' : 'none';

    showPage(0);
    console.log(`文档预览已切换为单页模式，共 ${totalPages} 页`);
}

function displayParsedContent(mdContent) {
    updateDocumentContent(mdContent);
}

function renderMarkdown(markdown, violations = []) {
    updateDocumentContent(markdown);

    setTimeout(() => {
        if (violations.length > 0) {
            applyViolationHighlights(violations);
        }
    }, 80);
}

// 在文件上传成功后的处理中调用
function handleDocumentParsed(data) {
    updateDocumentContent(data.md_content);

    document.getElementById('file-name').textContent = data.doc_name;
    document.getElementById('check-btn').disabled = false;
    document.getElementById('parse-status').style.display = 'none';
    document.getElementById('parsed-info').style.display = 'block';
    document.getElementById('parsed-doc-name').textContent = data.doc_name;
    document.getElementById('parsed-image-count').textContent = data.image_count;
    document.getElementById('parsed-images-dir').textContent =
        data.vl_special_layouts > 0
            ? `发现 ${data.vl_special_layouts} 个特殊布局`
            : '无特殊布局';
}

// 页面加载完成后初始化分页功能
document.addEventListener('DOMContentLoaded', () => {
    setupPaginationEvents();
    initPagination();
});

// 在需要更新高亮的地方触发事件
function triggerHighlightsUpdate() {
    if (checkResult?.violations?.length) {
        setTimeout(() => {
            applyViolationHighlights(checkResult.violations);
        }, 50);
    }
}

function toggleHighlight() {
    highlightEnabled = !highlightEnabled;
    const highlights = document.querySelectorAll('.violation-highlight');

    if (highlightEnabled) {
        toggleHighlightBtn.textContent = '关闭高亮';
        toggleHighlightBtn.classList.add('active');
        highlights.forEach((el) => {
            el.classList.add('highlight-visible');
            el.classList.remove('highlight-hidden');
        });
    } else {
        toggleHighlightBtn.textContent = '开启高亮';
        toggleHighlightBtn.classList.remove('active');
        highlights.forEach((el) => {
            el.classList.add('highlight-hidden');
            el.classList.remove('highlight-visible');
        });
    }

    triggerHighlightsUpdate();
}

// 从中间栏滚动到文档中的违规位置
function scrollToDocumentViolation(violationIndex) {
    const violation = checkResult?.violations?.[violationIndex];
    const targetPageNumber = Number(violation?.page_number);
    const targetPageIndex = Number.isNaN(targetPageNumber) ? -1 : getPageIndexByNumber(targetPageNumber);

    if (targetPageIndex !== -1 && targetPageIndex !== currentPageIndex) {
        showPage(targetPageIndex);
    }

    setTimeout(() => {
        const element = document.querySelector(`#page-content-container [data-violation-index="${violationIndex}"]`)
            || document.querySelector(`#document-content [data-violation-index="${violationIndex}"]`);

        if (element) {
            element.scrollIntoView({ behavior: 'smooth', block: 'center' });
            element.classList.add('flashing');
            setTimeout(() => element.classList.remove('flashing'), 2000);
        }
    }, 120);
}

// 点击右侧栏违规项 → 展开详情 + 中间栏滚动到对应位置
function handleViolationClick(index) {
    toggleViolation(index);
    scrollToDocumentViolation(index);
}

// 调试函数 - 在控制台执行可检查当前高亮状态
function debugHighlights() {
    console.log('当前高亮状态:');
    console.log('- 高亮开关:', highlightEnabled);
    console.log('- 违规数据:', checkResult?.violations?.length || 0);
    console.log('- 当前页面:', currentDocumentPages[currentPageIndex]?.pageNumber || '无');
    console.log('- 高亮元素:', document.querySelectorAll('.violation-highlight').length);
}

// 手动触发高亮
function forceApplyHighlights() {
    if (checkResult?.violations) {
        applyViolationHighlights(checkResult.violations);
        console.log('手动触发了高亮应用');
    }
}
