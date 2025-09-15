/**
 * 模板渲染和预览功能模块
 */

// 模板渲染核心函数
function renderTemplate(template, customerData) {
  if (!template || !customerData) return '';
  
  // 处理普通文本模板
  let rendered = template;
  
  // 处理方括号格式的字段变量 [[字段名]]
  rendered = rendered.replace(/\[\[([^\]]+)\]\]/g, (match, fieldName) => {
    const trimmedField = fieldName.trim();
    if (customerData.hasOwnProperty(trimmedField)) {
      const value = customerData[trimmedField];
      return value !== undefined && value !== null ? String(value) : '';
    }
    return '[暂无数据]';
  });
  
  return rendered;
}

// 更新预览内容
function updatePreview() {
  if (!appState.currentCustomer || !appState.template) {
    $('previewResult').innerHTML = '<div class="preview-empty">请选择客户并编辑模板</div>';
    return;
  }
  
  try {
    const rendered = renderTemplate(appState.template, appState.currentCustomer);
    // 将换行符转换为HTML换行标签，并转义HTML特殊字符
    const escapedText = rendered
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/\n/g, '<br>');
    $('previewResult').innerHTML = escapedText;
  } catch (error) {
    $('previewResult').innerHTML = `<div style="color: var(--danger);">模板渲染错误: ${error.message}</div>`;
  }
}

// 历史话术记录功能
function showHistoryModal() {
  $('historyModal').classList.add('show');
  updateHistoryList();
}

function closeHistoryModal() {
  $('historyModal').classList.remove('show');
  $('historySearchInput').value = '';
}

function addToHistory(customerData, template, renderedScript) {
  const historyItem = {
    id: Date.now().toString(),
    customer: getCustomerDisplayName(customerData),
    customerData: { ...customerData },
    template: template,
    renderedScript: renderedScript,
    time: new Date().toLocaleString('zh-CN')
  };
  
  // 添加到历史记录开头（最新的在前）
  appState.scriptHistory.unshift(historyItem);
  
  // 限制历史记录数量（最多100条）
  if (appState.scriptHistory.length > 100) {
    appState.scriptHistory = appState.scriptHistory.slice(0, 100);
  }
  
  saveHistoryToStorage();
}

function updateHistoryList() {
  const searchTerm = $('historySearchInput').value.toLowerCase().trim();
  const historyList = $('historyList');
  
  let filteredHistory = appState.scriptHistory;
  
  // 搜索过滤
  if (searchTerm) {
    filteredHistory = appState.scriptHistory.filter(item => 
      item.customer.toLowerCase().includes(searchTerm) ||
      item.renderedScript.toLowerCase().includes(searchTerm) ||
      item.template.toLowerCase().includes(searchTerm)
    );
  }
  
  if (filteredHistory.length === 0) {
    historyList.innerHTML = `
      <div style="text-align: center; color: var(--muted); padding: 20px;">
        ${searchTerm ? '没有找到相关记录' : '暂无使用记录'}
      </div>
    `;
    return;
  }
  
  historyList.innerHTML = '';
  
  filteredHistory.forEach(item => {
    const historyItem = document.createElement('div');
    historyItem.className = 'history-item';
    historyItem.innerHTML = `
      <div class="history-customer">客户：${item.customer}</div>
      <div class="history-preview">${item.renderedScript.substring(0, 120)}${item.renderedScript.length > 120 ? '...' : ''}</div>
      <div class="history-time">使用时间：${item.time}</div>
      <div class="history-actions">
        <button class="history-btn" onclick="useHistoryScript('${item.id}')">使用话术</button>
        <button class="history-btn" onclick="useHistoryTemplate('${item.id}')">使用模板</button>
        <button class="history-btn" onclick="previewHistory('${item.id}')">预览</button>
        <button class="history-btn delete" onclick="deleteHistory('${item.id}')">删除</button>
      </div>
    `;
    
    historyList.appendChild(historyItem);
  });
}

function useHistoryScript(historyId) {
  const historyItem = appState.scriptHistory.find(h => h.id === historyId);
  if (historyItem) {
    // 复制话术到剪贴板
    navigator.clipboard.writeText(historyItem.renderedScript).then(() => {
      updateStatus(`已复制历史话术到剪贴板`, 'success');
      closeHistoryModal();
    }).catch(() => {
      alert('复制失败，请手动复制内容');
    });
  }
}

function useHistoryTemplate(historyId) {
  const historyItem = appState.scriptHistory.find(h => h.id === historyId);
  if (historyItem) {
    $('templateEditor').value = historyItem.template;
    if (typeof handleTextChange === 'function') {
      handleTextChange();
    }
    closeHistoryModal();
    updateStatus(`已加载历史模板`, 'success');
  }
}

function previewHistory(historyId) {
  const historyItem = appState.scriptHistory.find(h => h.id === historyId);
  if (historyItem) {
    const previewWindow = window.open('', '_blank', 'width=600,height=400');
    previewWindow.document.write(`
      <html>
        <head>
          <title>话术预览 - ${historyItem.customer}</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; }
            .header { border-bottom: 1px solid #ccc; padding-bottom: 10px; margin-bottom: 20px; }
            .content { white-space: pre-wrap; }
          </style>
        </head>
        <body>
          <div class="header">
            <h3>客户：${historyItem.customer}</h3>
            <p>使用时间：${historyItem.time}</p>
          </div>
          <div class="content">${historyItem.renderedScript}</div>
        </body>
      </html>
    `);
    previewWindow.document.close();
  }
}

function deleteHistory(historyId) {
  if (confirm('确定要删除这条历史记录吗？')) {
    appState.scriptHistory = appState.scriptHistory.filter(h => h.id !== historyId);
    saveHistoryToStorage();
    updateHistoryList();
    updateStatus('历史记录已删除', 'success');
  }
}

function clearAllHistory() {
  if (confirm('确定要清空所有历史记录吗？此操作不可撤销。')) {
    appState.scriptHistory = [];
    saveHistoryToStorage();
    updateHistoryList();
    updateStatus('所有历史记录已清空', 'success');
  }
}

function filterHistoryList() {
  updateHistoryList();
}

function saveHistoryToStorage() {
  try {
    localStorage.setItem('scriptEditor_history', JSON.stringify(appState.scriptHistory));
  } catch (error) {
    console.error('保存历史记录失败:', error);
  }
}

function loadHistoryFromStorage() {
  try {
    const saved = localStorage.getItem('scriptEditor_history');
    if (saved) {
      appState.scriptHistory = JSON.parse(saved);
    }
  } catch (error) {
    console.error('加载历史记录失败:', error);
    appState.scriptHistory = [];
  }
}

// 计算字段相关功能
function showCalcFieldModal() {
  if (appState.fields.length === 0) {
    alert('请先上传数据文件');
    return;
  }
  $('calcFieldModal').classList.add('show');
  $('calcFieldNameInput').value = '';
  $('calcFormulaInput').value = '';
  updateCalcFieldSelector();
  updateCalcPreview();
  
  // 初始化计算字段模态框（包括字段保护功能）
  initCalcFieldModal();
}

function closeCalcFieldModal() {
  $('calcFieldModal').classList.remove('show');
}

function initCalcFieldModal() {
  // 初始化计算字段模态框
  const calcFormulaInput = $('calcFormulaInput');
  const calcFieldNameInput = $('calcFieldNameInput');
  
  // 移除可能已存在的事件监听器，防止重复绑定
  calcFormulaInput.removeEventListener('keydown', handleCalcFormulaKeydown);
  calcFormulaInput.removeEventListener('input', updateCalcPreview);
  calcFieldNameInput.removeEventListener('input', updateCalcPreview);
  
  // 绑定键盘事件以保护字段变量
  calcFormulaInput.addEventListener('keydown', handleCalcFormulaKeydown);
  
  // 绑定输入事件以更新预览（添加防抖）
  let previewTimeout;
  const debouncedUpdatePreview = function() {
    // 清除之前的定时器
    clearTimeout(previewTimeout);
    // 设置新的定时器，延迟300毫秒执行
    previewTimeout = setTimeout(updateCalcPreview, 300);
  };
  
  calcFormulaInput.addEventListener('input', debouncedUpdatePreview);
  calcFieldNameInput.addEventListener('input', debouncedUpdatePreview);
  
  // 更新字段选择器
  updateCalcFieldSelector();
}

// 处理计算公式输入框的键盘事件
function handleCalcFormulaKeydown(e) {
  const editor = e.target;
  const cursorPos = editor.selectionStart;
  
  // 支持 Tab 键缩进
  if (e.key === 'Tab') {
    e.preventDefault();
    const start = editor.selectionStart;
    const end = editor.selectionEnd;
    
    // 插入制表符
    editor.value = editor.value.substring(0, start) + '  ' + editor.value.substring(end);
    editor.selectionStart = editor.selectionEnd = start + 2;
    
    updateCalcPreview();
    return;
  }
  
  // 字段保护机制：防止部分删除字段变量
  if (e.key === 'Backspace' || e.key === 'Delete') {
    const result = protectCalcFieldVariables(editor, e.key, cursorPos);
    if (result.preventDefault) {
      e.preventDefault();
      if (result.newValue !== undefined) {
        editor.value = result.newValue;
        editor.selectionStart = editor.selectionEnd = result.newCursor;
        updateCalcPreview();
      }
    }
  }
}

// 字段保护机制：防止部分删除计算公式中的字段变量
function protectCalcFieldVariables(editor, key, cursorPos) {
  const text = editor.value;
  
  // 限制文本长度以防止性能问题
  if (text.length > 10000) {
    return { preventDefault: false };
  }
  
  // 查找所有字段变量
  const fieldRegex = /\[\[([^\]]+)\]\]/g;
  let match;
  const fields = [];
  let matchCount = 0;
  
  // 限制匹配次数以防止性能问题
  while ((match = fieldRegex.exec(text)) !== null && matchCount < 100) {
    fields.push({
      start: match.index,
      end: match.index + match[0].length,
      content: match[0]
    });
    matchCount++;
  }
  
  // 检查光标位置是否在字段变量内部或边界
  for (const field of fields) {
    // 检查光标是否在字段范围内（包括边界）
    if (cursorPos >= field.start && cursorPos <= field.end) {
      // 任何在字段范围内的删除操作都应该删除整个字段
      if (key === 'Backspace' || key === 'Delete') {
        return {
          preventDefault: true,
          newValue: text.substring(0, field.start) + text.substring(field.end),
          newCursor: field.start
        };
      }
    }
  }
  
  return { preventDefault: false };
}

function updateCalcFieldSelector() {
  const selector = $('calcFieldSelector');
  
  if (appState.fields.length === 0 && appState.calcFields.length === 0) {
    selector.innerHTML = '<div style="text-align: center; color: var(--muted); font-size: 11px; grid-column: 1 / -1;">请先上传数据文件</div>';
    return;
  }
  
  selector.innerHTML = '';
  
  // 添加原始字段
  appState.fields.forEach(field => {
    const fieldTag = document.createElement('button');
    fieldTag.className = 'field-tag';
    fieldTag.textContent = field.name;
    fieldTag.title = `点击插入字段: ${field.name}`;
    fieldTag.onclick = () => insertFieldIntoFormula(field.name);
    selector.appendChild(fieldTag);
  });
  
  // 添加计算字段
  appState.calcFields.forEach(calcField => {
    const fieldTag = document.createElement('button');
    fieldTag.className = 'field-tag';
    fieldTag.style.background = '#f59e0b'; // 计算字段用不同颜色
    fieldTag.textContent = `🧮 ${calcField.name}`;
    fieldTag.title = `点击插入计算字段: ${calcField.name}`;
    fieldTag.onclick = () => insertFieldIntoFormula(calcField.name);
    selector.appendChild(fieldTag);
  });
}

function insertFieldIntoFormula(fieldName) {
  const textarea = $('calcFormulaInput');
  const cursorPos = textarea.selectionStart;
  const textBefore = textarea.value.substring(0, cursorPos);
  const textAfter = textarea.value.substring(textarea.selectionEnd);
  const fieldReference = `[[${fieldName}]]`;
  
  textarea.value = textBefore + fieldReference + textAfter;
  
  // 设置光标位置到插入内容后面
  const newCursorPos = cursorPos + fieldReference.length;
  textarea.focus();
  textarea.setSelectionRange(newCursorPos, newCursorPos);
  
  // 更新预览
  updateCalcPreview();
}

// 插入函数模板
function insertFunction(template) {
  const textarea = $('calcFormulaInput');
  const cursorPos = textarea.selectionStart;
  const textBefore = textarea.value.substring(0, cursorPos);
  const textAfter = textarea.value.substring(textarea.selectionEnd);
  
  textarea.value = textBefore + template + textAfter;
  
  // 设置光标位置
  let newCursorPos;
  if (template === ' ?  : ') {
    // 对于条件表达式，将光标放在问号后面
    newCursorPos = cursorPos + 2;
  } else if (template.includes('()')) {
    // 对于函数，将光标放在括号内
    newCursorPos = cursorPos + template.indexOf(')');
  } else {
    newCursorPos = cursorPos + template.length;
  }
  
  textarea.focus();
  textarea.setSelectionRange(newCursorPos, newCursorPos);
  
  // 更新预览
  updateCalcPreview();
}

function updateCalcPreview() {
  const fieldName = $('calcFieldNameInput').value.trim();
  const formula = $('calcFormulaInput').value.trim();
  const previewDiv = $('calcPreview');
  
  // 检查输入是否完整
  if (!fieldName || !formula) {
    previewDiv.innerHTML = '<span style="color: var(--muted);">请输入字段名称和公式...</span>';
    $('addCalcFieldConfirmBtn').disabled = true;
    return;
  }
  
  // 限制输入长度以防止性能问题
  if (fieldName.length > 100 || formula.length > 1000) {
    previewDiv.innerHTML = '<span style="color: var(--warning);">输入内容过长</span>';
    $('addCalcFieldConfirmBtn').disabled = true;
    return;
  }
  
  // 检查字段名是否已存在
  if (appState.fields.some(f => f.name === fieldName) || appState.calcFields.some(f => f.name === fieldName)) {
    previewDiv.innerHTML = `<span style="color: var(--danger);">字段名称已存在，请使用其他名称</span>`;
    $('addCalcFieldConfirmBtn').disabled = true;
    return;
  }
  
  try {
    // 预览计算结果（使用第一条数据）
    if (appState.data.length > 0) {
      const sampleData = appState.data[0];
      const evaluationResult = evaluateFormulaWithSteps(formula, sampleData);
      
      // 限制步骤数量以防止过多的DOM操作
      const displaySteps = evaluationResult.steps.slice(0, 30);
      
      // 构建预览内容，包含计算过程和结果
      let previewContent = `
        <div style="color: var(--success); margin-bottom: 8px;">✓ 公式有效</div>
        <div style="font-size: 11px; color: var(--text); margin-bottom: 6px; font-weight: 500;">计算过程：</div>
        <div style="background: var(--bg-secondary); border: 1px solid var(--border-color); border-radius: 6px; padding: 10px; margin-bottom: 8px; font-family: 'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace; font-size: 12px; line-height: 1.6; max-height: 200px; overflow-y: auto;">
          ${displaySteps.map(step => `
            <div style="margin-bottom: 1px; display: flex; align-items: flex-start;">
              <span style="color: var(--accent); margin-right: 6px;">→</span>
              <span>${step}</span>
            </div>
          `).join('')}
        </div>
        <div style="font-size: 11px; color: var(--muted); padding-top: 4px; border-top: 1px solid var(--border-color);">
          示例结果（第1条数据）：<strong style="color: var(--text);">${evaluationResult.result}</strong>
        </div>
      `;
      
      previewDiv.innerHTML = previewContent;
      $('addCalcFieldConfirmBtn').disabled = false;
    } else {
      previewDiv.innerHTML = '<span style="color: var(--warning);">没有数据可供预览</span>';
      $('addCalcFieldConfirmBtn').disabled = false;
    }
  } catch (error) {
    console.error('计算预览错误:', error); // 添加错误日志
    previewDiv.innerHTML = `<span style="color: var(--danger);">公式错误：${error.message}</span>`;
    $('addCalcFieldConfirmBtn').disabled = true;
  }
}

function addCalculatedField() {
  const fieldName = $('calcFieldNameInput').value.trim();
  const formula = $('calcFormulaInput').value.trim();
  
  if (!fieldName || !formula) {
    alert('请输入字段名称和公式');
    return;
  }
  
  // 添加计算字段
  const calcField = {
    name: fieldName,
    formula: formula
  };
  
  appState.calcFields.push(calcField);
  
  // 应用计算字段到数据
  applyCalculatedFields();
  
  // 更新界面
  updateFieldsList();
  updateCalcFieldSelector();
  updateDataOverview($('fileName').textContent, appState.data.length, appState.fields.length + appState.calcFields.length);
  updateNameColumnSelector();
  
  closeCalcFieldModal();
  updateStatus(`计算字段「${fieldName}」已添加`, 'success');
}

function applyCalculatedFields() {
  appState.data.forEach(row => {
    appState.calcFields.forEach(calcField => {
      try {
        row[calcField.name] = evaluateFormula(calcField.formula, row);
      } catch (error) {
        row[calcField.name] = '计算错误';
        console.error(`计算字段 ${calcField.name} 错误:`, error);
      }
    });
  });
}

function evaluateFormula(formula, data) {
  // 替换字段引用
  let processedFormula = formula.replace(/\[\[([^\]]+)\]\]/g, (match, fieldName) => {
    const trimmedField = fieldName.trim();
    if (data.hasOwnProperty(trimmedField)) {
      const value = data[trimmedField];
      // 如果是数字，直接返回；如果是字符串，用引号包裹
      if (value === '' || value === null || value === undefined) {
        return '""';
      }
      
      // 尝试解析为数字
      const numValue = parseFloat(value);
      if (!isNaN(numValue) && isFinite(numValue)) {
        return numValue.toString();
      }
      
      // 否则作为字符串处理
      return '"' + String(value).replace(/"/g, '\\"') + '"';
    }
    return '""';
  });
  
  // 安全性检查：限制可用的函数和操作
  const allowedPatterns = [
    /^[\s\d+\-*/()."'\w\u4e00-\u9fa5><=!?:&|]+$/,  // 只允许基本操作符、数字、字符串、中文和英文
  ];
  
  if (!allowedPatterns.some(pattern => pattern.test(processedFormula))) {
    throw new Error('公式包含不允许的字符');
  }
  
  // 禁止危险函数
  const dangerousKeywords = ['eval', 'Function', 'setTimeout', 'setInterval', 'require', 'import', 'export'];
  if (dangerousKeywords.some(keyword => processedFormula.includes(keyword))) {
    throw new Error('公式包含禁止的函数');
  }
  
  try {
    // 使用 Function 构造函数安全执行
    const result = new Function('Math', `"use strict"; return (${processedFormula})`)(Math);
    return result;
  } catch (error) {
    throw new Error('公式执行错误: ' + error.message);
  }
}

// 带计算步骤的公式评估函数
function evaluateFormulaWithSteps(formula, data) {
  // 限制公式长度防止性能问题
  if (formula.length > 1000) {
    throw new Error('公式过长');
  }
  
  const steps = [];
  
  // 限制步骤数量防止过多输出
  const maxSteps = 50;
  
  // 初始公式
  steps.push(`原始公式: ${formula}`);
  
  // 替换字段引用并记录步骤
  let processedFormula = formula.replace(/\[\[([^\]]+)\]\]/g, (match, fieldName) => {
    // 限制步骤数量
    if (steps.length >= maxSteps) {
      return match; // 如果步骤过多，不再处理
    }
    
    const trimmedField = fieldName.trim();
    if (data.hasOwnProperty(trimmedField)) {
      const value = data[trimmedField];
      let displayValue, rawValue;
      
      // 如果是数字，直接返回；如果是字符串，用引号包裹
      if (value === '' || value === null || value === undefined) {
        displayValue = '""';
        rawValue = '""';
      } else {
        // 尝试解析为数字
        const numValue = parseFloat(value);
        if (!isNaN(numValue) && isFinite(numValue)) {
          displayValue = numValue.toString();
          rawValue = numValue.toString();
        } else {
          // 否则作为字符串处理
          const escapedValue = String(value).replace(/"/g, '\\"');
          displayValue = `"${escapedValue}"`;
          rawValue = `"${escapedValue}"`;
        }
      }
      
      // 记录替换步骤
      steps.push(`替换字段 [[${trimmedField}]] → ${displayValue}`);
      return rawValue;
    }
    
    steps.push(`字段 [[${fieldName}]] 未找到 → ""`);
    return '""';
  });
  
  // 限制步骤数量
  if (steps.length < maxSteps) {
    steps.push(`替换后公式: ${processedFormula}`);
  }
  
  // 安全性检查：限制可用的函数和操作
  const allowedPatterns = [
    /^[\s\d+\-*/()."'\w\u4e00-\u9fa5><=!?:&|]+$/,  // 只允许基本操作符、数字、字符串、中文和英文
  ];
  
  if (!allowedPatterns.some(pattern => pattern.test(processedFormula))) {
    throw new Error('公式包含不允许的字符');
  }
  
  // 禁止危险函数
  const dangerousKeywords = ['eval', 'Function', 'setTimeout', 'setInterval', 'require', 'import', 'export'];
  if (dangerousKeywords.some(keyword => processedFormula.includes(keyword))) {
    throw new Error('公式包含禁止的函数');
  }
  
  try {
    // 对于简单表达式，尝试分步计算
    if (steps.length < maxSteps) {
      steps.push('开始计算...');
    }
    
    // 如果是简单算术表达式，尝试分步显示
    if (/^[0-9+\-*/.() ]+$/.test(processedFormula.replace(/\s/g, ''))) {
      try {
        // 尝试分步计算（仅适用于简单算术表达式）
        const result = evaluateArithmeticWithSteps(processedFormula, steps);
        if (steps.length < maxSteps) {
          steps.push(`计算完成`);
        }
        return {
          result: result,
          steps: steps.slice(0, maxSteps) // 限制步骤数量
        };
      } catch (e) {
        // 如果分步计算失败，回退到直接计算
        if (steps.length < maxSteps) {
          steps.push('分步计算失败，使用直接计算');
        }
      }
    }
    
    // 使用 Function 构造函数安全执行
    const result = new Function('Math', `"use strict"; return (${processedFormula})`)(Math);
    if (steps.length < maxSteps) {
      steps.push(`计算完成: ${result}`);
    }
    
    return {
      result: result,
      steps: steps.slice(0, maxSteps) // 限制步骤数量
    };
  } catch (error) {
    if (steps.length < maxSteps) {
      steps.push(`计算错误: ${error.message}`);
    }
    throw new Error('公式执行错误: ' + error.message);
  }
}

// 简单算术表达式的分步计算
function evaluateArithmeticWithSteps(expression, steps) {
  // 移除空格
  let expr = expression.replace(/\s+/g, '');
  if (expr !== expression) {
    steps.push(`简化表达式: ${expr}`);
  }
  
  // 简单的分步计算逻辑（支持基本的 + - * / 和括号）
  // 这是一个简化的实现，仅用于演示目的
  
  // 处理括号（从最内层开始）
  let iteration = 0;
  const maxIterations = 20; // 限制最大迭代次数防止无限循环
  while (expr.includes('(') && iteration < maxIterations) {
    iteration++;
    // 查找最内层的括号
    let innermostStart = -1;
    let innermostEnd = -1;
    for (let i = 0; i < expr.length; i++) {
      if (expr[i] === '(') {
        innermostStart = i;
      } else if (expr[i] === ')') {
        innermostEnd = i;
        break;
      }
    }
    
    // 如果找到了括号对
    if (innermostStart !== -1 && innermostEnd !== -1 && innermostEnd > innermostStart) {
      const innerExpr = expr.substring(innermostStart + 1, innermostEnd);
      const result = evaluateSimpleExpression(innerExpr);
      steps.push(`计算括号: (${innerExpr}) = ${result}`);
      expr = expr.substring(0, innermostStart) + result.toString() + expr.substring(innermostEnd + 1);
      steps.push(`处理后表达式: ${expr}`);
    } else {
      // 如果没有找到有效的括号对，跳出循环
      break;
    }
  }
  
  // 处理剩余表达式
  const finalResult = evaluateSimpleExpression(expr);
  steps.push(`最终计算: ${expr} = ${finalResult}`);
  
  return finalResult;
}

// 简单表达式计算（不包含括号）
function evaluateSimpleExpression(expr) {
  // 验证表达式是否为空或无效
  if (!expr || expr.trim() === '') {
    return 0;
  }
  
  // 先处理乘法和除法（从左到右）
  let changed = true;
  while (changed && (expr.includes('*') || expr.includes('/'))) {
    changed = false;
    expr = expr.replace(/(-?\d*\.?\d+)([*/])(-?\d*\.?\d+)/, (match, a, op, b) => {
      const numA = parseFloat(a);
      const numB = parseFloat(b);
      let result;
      if (op === '*') {
        result = numA * numB;
      } else if (op === '/') {
        // 处理除零错误
        if (numB === 0) {
          throw new Error('除零错误');
        }
        result = numA / numB;
      }
      changed = true;
      return result.toString();
    });
  }
  
  // 再处理加法和减法（从左到右）
  changed = true;
  while (changed && (expr.includes('+') || (expr.match(/-/g) || []).length > (expr.startsWith('-') ? 1 : 0))) {
    changed = false;
    // 处理减法时需要考虑负数的情况
    expr = expr.replace(/(-?\d*\.?\d+)([+\-])(-?\d*\.?\d+)/, (match, a, op, b) => {
      const numA = parseFloat(a);
      const numB = parseFloat(b);
      let result;
      if (op === '+') {
        result = numA + numB;
      } else if (op === '-') {
        result = numA - numB;
      }
      changed = true;
      return result.toString();
    });
  }
  
  const result = parseFloat(expr);
  if (isNaN(result)) {
    throw new Error('无效的表达式');
  }
  return result;
}

// 导出功能
function showExportModal() {
  // 创建模态框
  const modal = document.createElement('div');
  modal.style.cssText = `
    position: fixed; top: 0; left: 0; width: 100%; height: 100%;
    background: rgba(0, 0, 0, 0.5); z-index: 10001;
    display: flex; align-items: center; justify-content: center;
  `;
  
  modal.innerHTML = `
    <div style="
      background: var(--panel); border-radius: 12px; padding: 24px;
      box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15); max-width: 400px; width: 90%;
    ">
      <h3 style="margin: 0 0 16px 0; font-size: 18px; color: var(--text); text-align: center;">
        📤 选择导出格式
      </h3>
      <p style="margin: 0 0 20px 0; color: var(--muted); text-align: center; line-height: 1.5;">
        请选择您需要的导出格式：
      </p>
      
      <div style="display: flex; flex-direction: column; gap: 12px;">
        <button onclick="exportAsExcel(); closeExportModal(this)" 
                style="
                  padding: 12px 16px; background: var(--success); color: white;
                  border: none; border-radius: 8px; cursor: pointer; font-size: 14px;
                  transition: all 0.2s; display: flex; align-items: center; gap: 8px;
                "
                onmouseover="this.style.background='#059669'"
                onmouseout="this.style.background='var(--success)'">
          <span>📈</span>
          <div style="text-align: left;">
            <div style="font-weight: 600;">导出Excel表格</div>
            <div style="font-size: 12px; opacity: 0.9;">结构化数据，包含完整信息，适合分析</div>
          </div>
        </button>
        
        <button onclick="exportAsText(); closeExportModal(this)" 
                style="
                  padding: 12px 16px; background: var(--accent); color: white;
                  border: none; border-radius: 8px; cursor: pointer; font-size: 14px;
                  transition: all 0.2s; display: flex; align-items: center; gap: 8px;
                "
                onmouseover="this.style.background='#1d4ed8'"
                onmouseout="this.style.background='var(--accent)'">
          <span>📝</span>
          <div style="text-align: left;">
            <div style="font-weight: 600;">导出纯文本</div>
            <div style="font-size: 12px; opacity: 0.9;">传统文本格式，适合直接使用</div>
          </div>
        </button>
        
        <button onclick="closeExportModal(this)" 
                style="
                  padding: 10px 16px; background: var(--panel); color: var(--muted);
                  border: 1px solid var(--border); border-radius: 8px; cursor: pointer;
                  font-size: 14px; transition: all 0.2s;
                "
                onmouseover="this.style.background='#f3f4f6'"
                onmouseout="this.style.background='var(--panel)'">
          取消导出
        </button>
      </div>
    </div>
  `;
  
  // 点击背景关闭
  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      document.body.removeChild(modal);
    }
  });
  
  document.body.appendChild(modal);
}

function closeExportModal(button) {
  // 找到最外层的模态框元素
  let modal = button;
  
  // 向上遍历找到最外层的div（模态框）
  while (modal && modal !== document.body) {
    if (modal.style && modal.style.position === 'fixed') {
      // 找到了模态框
      document.body.removeChild(modal);
      return;
    }
    modal = modal.parentElement;
  }
  
  // 如果上面的方法失败，使用备用方法
  const allModals = document.querySelectorAll('div[style*="position: fixed"]');
  allModals.forEach(m => {
    if (m.style.zIndex === '10001') {
      document.body.removeChild(m);
    }
  });
}

function exportAsExcel() {
  try {
    // 创建表格数据
    const exportData = [];
    
    // 表头：包含所有原始字段 + 计算字段 + 生成话术
    const headers = [
      ...appState.fields.map(f => f.name),
      ...appState.calcFields.map(f => `🧮 ${f.name}`),
      '📝 生成话术'
    ];
    exportData.push(headers);
    
    // 数据行：每个客户的完整信息 + 生成的话术
    appState.data.forEach((customer, index) => {
      const row = [];
      
      // 添加原始字段数据
      appState.fields.forEach(field => {
        row.push(customer[field.name] || '');
      });
      
      // 添加计算字段数据
      appState.calcFields.forEach(calcField => {
        row.push(customer[calcField.name] || '');
      });
      
      // 添加生成的话术
      const renderedScript = renderTemplate(appState.template, customer);
      row.push(renderedScript);
      
      exportData.push(row);
    });
    
    // 生成CSV内容（Excel可以正确读取CSV）
    const csvContent = exportData.map(row => 
      row.map(cell => {
        // 处理包含逗号、换行符或引号的单元格
        const cellStr = String(cell || '');
        if (cellStr.includes(',') || cellStr.includes('\n') || cellStr.includes('"')) {
          return '"' + cellStr.replace(/"/g, '""') + '"';
        }
        return cellStr;
      }).join(',')
    ).join('\n');
    
    // 添加BOM以支持中文
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8' });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `批量话术导出_${new Date().toISOString().slice(0, 10)}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    updateStatus(`已导出Excel表格，包含 ${appState.data.length} 条记录`, 'success');
  } catch (error) {
    alert('导出失败: ' + error.message);
  }
}

function exportAsText() {
  try {
    const results = appState.data.map((customer, index) => {
      const customerName = getCustomerDisplayName(customer);
      const rendered = renderTemplate(appState.template, customer);
      return `=== 客户 ${index + 1}: ${customerName} ===\n${rendered}\n\n`;
    }).join('');
    
    const blob = new Blob([results], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `批量话术导出_${new Date().toISOString().slice(0, 10)}.txt`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    updateStatus(`已导出文本格式，包含 ${appState.data.length} 条话术`, 'success');
  } catch (error) {
    alert('导出失败: ' + error.message);
  }
}