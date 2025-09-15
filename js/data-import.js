/**
 * 数据导入和处理功能模块
 */

// 导入配置相关变量
let currentImportFile = null;
let currentWorkbook = null;
let currentFileData = null;

// 文件处理
window.handleFileUpload = function handleFileUpload(event) {
  console.log('文件上传事件触发');
  const file = event.target.files[0];
  if (!file) {
    console.log('没有选择文件');
    return;
  }
  
  console.log(`选择的文件: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`);
  currentImportFile = file;
  
  try {
    showImportConfigModal(file);
    console.log('文件上传处理完成，显示导入配置模态框');
  } catch (error) {
    const errorMsg = '文件上传处理失败: ' + (error.message || error);
    console.error(errorMsg);
    updateStatus(errorMsg, 'error');
  }
  
  // 不要重置文件输入的值，这可能会中断文件处理
  // event.target.value = '';
}

// 显示导入配置模态框
async function showImportConfigModal(file) {
  try {
    console.log('开始显示导入配置模态框，文件:', file.name);
    
    // 显示文件信息
    updateFileInfo(file);
    
    // 显示模态框
    const modal = document.getElementById('importConfigModal');
    if (!modal) {
      const errorMsg = '未找到导入配置模态框元素';
      console.error(errorMsg);
      return;
    }
    
    modal.classList.add('show');
    
    // 确保显示
    modal.style.display = 'flex';
    
    updateImportStatus('正在解析文件...', 'info');
    
    // 解析文件并准备预览
    await prepareFileForImport(file);
    
    // 更新预览
    updateDataPreview();
    
    // 更新表头行选择器
    updateHeaderRowSelect();
    
    updateImportStatus('文件已加载，请配置导入选项', 'info');
    console.log('文件解析成功，准备导入');
    
    // 启用确认导入按钮
    const confirmBtn = document.getElementById('confirmImportBtn');
    if (confirmBtn) {
      confirmBtn.disabled = false;
    }
  } catch (error) {
    const errorMsg = '导入配置错误: ' + (error.message || error);
    if (typeof logger !== 'undefined' && logger.error) {
      logger.error(errorMsg);
      logger.operation('导入失败：' + errorMsg);
    } else {
      console.error('导入配置错误:', error);
    }
    updateImportStatus('文件解析失败: ' + error.message, 'error');
    
    // 确保确认导入按钮在错误状态下被禁用
    const confirmBtn = document.getElementById('confirmImportBtn');
    if (confirmBtn) {
      confirmBtn.disabled = true;
    }
    
    // 提供一些常见问题的解决建议
    if (error.message.includes('编码') || error.message.includes('乱码')) {
      updateImportStatus('文件编码问题，请尝试将文件保存为UTF-8编码格式', 'error');
    }
  }
}

function updateFileInfo(file) {
  console.log('更新文件信息:', file);
  const nameElement = document.getElementById('importFileName');
  const sizeElement = document.getElementById('importFileSize');
  const typeElement = document.getElementById('importFileType');
  
  if (nameElement) nameElement.textContent = file.name;
  if (sizeElement) sizeElement.textContent = `文件大小: ${(file.size / 1024).toFixed(1)} KB`;
  
  const extension = file.name.toLowerCase().split('.').pop();
  const typeMap = {
    'csv': 'CSV 逗号分隔文件',
    'tsv': 'TSV 制表符分隔文件', 
    'xlsx': 'Excel 工作簿文件',
    'xls': 'Excel 兼容文件'
  };
  
  if (typeElement) typeElement.textContent = `文件类型: ${typeMap[extension] || '未知格式'}`;
}

async function prepareFileForImport(file) {
  console.log('准备导入文件:', file);
  const fileName = file.name.toLowerCase();
  
  try {
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      console.log('处理Excel文件');
      await prepareExcelFile(file);
    } else if (fileName.endsWith('.csv') || fileName.endsWith('.tsv')) {
      console.log('处理文本文件');
      await prepareTextFile(file);
    } else {
      throw new Error('不支持的文件格式，请使用 CSV、TSV 或 Excel 文件');
    }
    
    // 确保currentFileData存在且不为空
    console.log('文件数据:', currentFileData);
    if (!currentFileData || currentFileData.length === 0) {
      throw new Error('文件中没有解析到有效数据');
    }
    
    console.log('文件准备完成，数据行数:', currentFileData.length);
  } catch (error) {
    console.error('文件准备过程中出错:', error);
    throw error; // 重新抛出错误以便上层处理
  }
}

async function prepareExcelFile(file) {
  console.log('开始处理Excel文件:', file.name);
  
  if (typeof XLSX === 'undefined') {
    throw new Error('Excel支持库未加载');
  }
  
  try {
    const arrayBuffer = await file.arrayBuffer();
    console.log('文件缓冲区大小:', arrayBuffer.byteLength);
    
    currentWorkbook = XLSX.read(arrayBuffer, { type: 'array' });
    console.log(`工作簿已加载，包含 ${currentWorkbook.SheetNames.length} 个工作表`);
    
    // 更新工作表选择器
const worksheetSelect = document.getElementById('worksheetSelect');
if (worksheetSelect) {
  worksheetSelect.innerHTML = '';
}

// 确保currentWorkbook.SheetNames存在且不为空
if (!currentWorkbook || !currentWorkbook.SheetNames || currentWorkbook.SheetNames.length === 0) {
  throw new Error('Excel文件中没有找到有效的工作表');
}

console.log(`工作表数量: ${currentWorkbook.SheetNames.length}`);

// 如果有多个工作表，显示选择器
if (currentWorkbook.SheetNames.length > 1 && worksheetSelect) {
  try {
    currentWorkbook.SheetNames.forEach((name, index) => {
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `第${index + 1}个工作表 (${name})`;
      worksheetSelect.appendChild(option);
    });
    
    document.getElementById('worksheetGroup').style.display = 'block';
    
    if (typeof logger !== 'undefined' && logger.log) {
      logger.log('已加载多个工作表选项');
    }
    
    // 更新表头行选择
    updateHeaderRowSelect();
  } catch (error) {
    const errorMsg = '更新工作表选择器失败: ' + (error.message || error);
    if (typeof logger !== 'undefined' && logger.error) {
      logger.error(errorMsg);
      logger.operation(errorMsg);
    } else {
      console.error(errorMsg);
    }
  }
} else {
  // 只有一个工作表，隐藏选择器
  const worksheetGroup = document.getElementById('worksheetGroup');
  if (worksheetGroup) {
    worksheetGroup.style.display = 'none';
  }
  
  // 即使只有一个工作表，也需要更新表头行选择
  if (currentWorkbook && currentWorkbook.SheetNames && currentWorkbook.SheetNames.length > 0) {
    updateHeaderRowSelect();
  }
}
    
    // 默认选择第一个工作表
    const worksheet = currentWorkbook.Sheets[currentWorkbook.SheetNames[0]];
    console.log('第一个工作表:', worksheet);
    currentFileData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log('解析后的数据:', currentFileData);
    
    // 确保currentFileData存在且不为空
    if (!currentFileData || currentFileData.length === 0) {
      throw new Error('Excel工作表中没有找到有效数据');
    }
  } catch (error) {
    console.error('Excel文件处理错误:', error);
    throw new Error('Excel文件处理失败: ' + error.message);
  }
}

async function prepareTextFile(file) {
  console.log('开始处理文本文件:', file);
  try {
    const text = await readFileWithProperEncoding(file);
    console.log('读取的文本内容长度:', text.length);
    // 处理BOM标记
    const cleanText = text.replace(/^\uFEFF/, '').trim();
    console.log('清理后的文本长度:', cleanText.length);
    if (!cleanText) {
      throw new Error('文件内容为空');
    }
    
    const lines = cleanText.split(/\r?\n/);
    console.log('行数:', lines.length);
    const delimiter = file.name.toLowerCase().includes('.tsv') ? '\t' : ',';
    console.log('分隔符:', delimiter);
    
    currentFileData = parseCSVLines(lines, delimiter);
    console.log('解析后的数据:', currentFileData);
    
    // 确保currentFileData不为空
    if (!currentFileData || currentFileData.length === 0) {
      throw new Error('文件中没有解析到有效数据');
    }
    
    const worksheetGroup = document.getElementById('worksheetGroup');
    if (worksheetGroup) {
      worksheetGroup.style.display = 'none';
    }
  } catch (error) {
    console.error('文本文件处理错误:', error);
    throw new Error('文本文件处理失败: ' + error.message);
  }
}

// 使用正确的编码读取文件
async function readFileWithProperEncoding(file) {
  console.log('开始读取文件编码:', file);
  // 尝试不同的编码格式
  const encodings = ['UTF-8', 'GBK', 'GB2312', 'Big5'];
  
  for (const encoding of encodings) {
    try {
      const text = await readFileWithEncoding(file, encoding);
      
      // 检查是否有乱码字符（替换字符）并且文本不为空
      if (text && !text.includes('\uFFFD') && text.trim().length > 0) {
        console.log(`成功使用 ${encoding} 编码读取文件`);
        return text;
      }
    } catch (error) {
      console.warn(`使用 ${encoding} 编码读取文件失败:`, error);
    }
  }
  
  // 如果所有编码都失败，使用默认UTF-8
  console.warn('所有编码尝试失败，使用默认UTF-8编码');
  return await readFileWithEncoding(file, 'UTF-8');
}

// 使用指定编码读取文件
function readFileWithEncoding(file, encoding) {
  console.log(`使用 ${encoding} 编码读取文件`);
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = function(e) {
      console.log('文件读取成功');
      resolve(e.target.result);
    };
    
    reader.onerror = function() {
      console.error(`使用 ${encoding} 编码读取文件失败`);
      reject(new Error(`使用 ${encoding} 编码读取文件失败`));
    };
    
    // 使用指定编码读取
    reader.readAsText(file, encoding);
  });
}

function parseCSVLines(lines, delimiter = ',') {
  console.log('解析CSV行，行数:', lines.length, '分隔符:', delimiter);
  const result = lines.map((line, index) => {
    // 添加日志以调试问题
    if (index < 5) {
      console.log(`处理第${index}行:`, line);
    }
    
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          // 处理转义的双引号
          current += '"';
          i++; // 跳过下一个引号
        } else {
          // 切换引号状态
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        // 在引号外遇到分隔符
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    
    result.push(current);
    return result.map(cell => cell.trim());
  });
  
  console.log('解析完成，结果行数:', result.length);
  
  // 只过滤掉完全空的行（所有单元格都为空或只包含空白字符）
  const filtered = result.filter((row, index) => {
    const keep = row.some(cell => cell && cell.toString().trim());
    if (index < 5) {
      console.log(`过滤第${index}行:`, row, '保留:', keep);
    }
    return keep;
  });
  
  console.log('过滤后行数:', filtered.length);
  return filtered;
}

function updateDataPreview() {
  console.log('更新数据预览，当前数据:', currentFileData);
  if (!currentFileData || currentFileData.length === 0) {
    const previewElement = document.getElementById('dataPreview');
    if (previewElement) {
      previewElement.innerHTML = '<div style="color: var(--muted); font-style: italic;">无数据</div>';
    }
    return;
  }
  
  const headerRowSelect = document.getElementById('headerRowSelect');
  const headerRow = headerRowSelect ? parseInt(headerRowSelect.value) - 1 : 0;
  const previewContainer = document.getElementById('dataPreview');
  
  if (!previewContainer) {
    console.error('未找到数据预览容器');
    return;
  }
  
  // 创建表格显示更好的中文支持
  const table = document.createElement('table');
  table.style.cssText = `
    width: 100%;
    border-collapse: collapse;
    font-family: 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
    font-size: 12px;
    white-space: nowrap;
  `;
  
  // 只显示前8行数据
  const previewData = currentFileData.slice(0, 8);
  
  previewData.forEach((row, index) => {
    const tr = document.createElement('tr');
    const isHeader = index === headerRow;
    
    if (isHeader) {
      tr.style.backgroundColor = '#e3f2fd';
      tr.style.fontWeight = 'bold';
    }
    
    // 行号列
    const lineNumberTd = document.createElement('td');
    lineNumberTd.textContent = (index + 1).toString().padStart(2, '0') + (isHeader ? '*' : ' ');
    lineNumberTd.style.cssText = `
      padding: 4px 8px;
      border: 1px solid #ddd;
      background-color: #f5f5f5;
      font-family: monospace;
      min-width: 30px;
      text-align: center;
    `;
    tr.appendChild(lineNumberTd);
    
    // 数据列
    row.forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      // 确保正确显示中文字符
      td.textContent = cell ? cell.toString() : '';
      td.style.cssText = `
        padding: 4px 8px;
        border: 1px solid #ddd;
        max-width: 120px;
        overflow: hidden;
        text-overflow: ellipsis;
      `;
      
      if (cellIndex < 5) { // 只显示前5列
        tr.appendChild(td);
      }
    });
    
    // 如果列数超过5列，显示省略号
    if (row.length > 5) {
      const ellipsisTd = document.createElement('td');
      ellipsisTd.textContent = `... (+${row.length - 5}列)`;
      ellipsisTd.style.cssText = `
        padding: 4px 8px;
        border: 1px solid #ddd;
        color: #666;
        font-style: italic;
      `;
      tr.appendChild(ellipsisTd);
    }
    
    table.appendChild(tr);
  });
  
  // 清空并添加新表格
  previewContainer.innerHTML = '';
  previewContainer.appendChild(table);
  
  // 添加说明文字
  const note = document.createElement('div');
  note.style.cssText = 'margin-top: 8px; font-size: 11px; color: #666;';
  note.textContent = `显示前${Math.min(previewData.length, 8)}行数据，带*号的为指定的字段行`;
  previewContainer.appendChild(note);
}

function updateImportStatus(message, type = 'info') {
  console.log('更新导入状态:', message, '类型:', type);
  const statusBadge = document.getElementById('importStatusBadge');
  if (statusBadge) {
    statusBadge.textContent = message;
    statusBadge.className = `status ${type}`;
  }
  
  // 根据状态启用/禁用确认按钮
  const confirmBtn = document.getElementById('confirmImportBtn');
  if (confirmBtn) {
    // 只有在错误状态时禁用按钮，其他状态都启用
    if (type === 'error') {
      confirmBtn.disabled = true;
    } else {
      confirmBtn.disabled = false;
    }
  }
}

// 关闭导入配置模态框
window.closeImportConfigModal = function closeImportConfigModal() {
  console.log('关闭导入配置模态框');
  const modal = document.getElementById('importConfigModal');
  if (modal) {
    // 确保移除show类
    modal.classList.remove('show');
    
    // 作为额外保障，直接设置样式
    modal.style.display = 'none';
    
    console.log('模态框已隐藏');
  } else {
    console.error('未找到导入配置模态框元素');
  }
  currentImportFile = null;
  currentWorkbook = null;
  currentFileData = null;
  
  // 确保功能按钮状态更新
  setTimeout(() => {
    if (typeof enableDataDependentFeatures === 'function') {
      enableDataDependentFeatures();
    } else if (typeof window.enableDataDependentFeatures === 'function') {
      window.enableDataDependentFeatures();
    }
  }, 100);
}

// 工作表选择变化处理
window.handleWorksheetChange = function handleWorksheetChange() {
  console.log('工作表选择变化');
  if (!currentWorkbook) {
    console.log('没有当前工作簿');
    return;
  }
  
  const worksheetSelect = document.getElementById('worksheetSelect');
  if (!worksheetSelect) {
    console.log('未找到工作表选择器');
    return;
  }
  
  const selectedIndex = parseInt(worksheetSelect.value);
  const sheetName = currentWorkbook.SheetNames[selectedIndex];
  const worksheet = currentWorkbook.Sheets[sheetName];
  
  currentFileData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  console.log('切换工作表后数据:', currentFileData);
  updateDataPreview();
  updateImportStatus('工作表已切换，请检查数据预览', 'info');
}

// 字段行选择变化处理
window.handleHeaderRowChange = function handleHeaderRowChange() {
  console.log('字段行选择变化');
  updateDataPreview();
  
  const headerRowSelect = document.getElementById('headerRowSelect');
  if (headerRowSelect) {
    const headerRow = parseInt(headerRowSelect.value);
    updateImportStatus(`字段行已设置为第${headerRow}行`, 'info');
  }
}

// 确认导入
window.confirmImport = async function confirmImport() {
  if (typeof logger !== 'undefined' && logger.log) {
    logger.log('确认导入函数被调用');
    logger.operation('开始导入数据');
  } else {
    console.log('确认导入函数被调用');
  }
  
  if (!currentFileData || !Array.isArray(currentFileData)) {
    const errorMsg = '没有可导入的数据或数据格式不正确';
    if (typeof logger !== 'undefined' && logger.error) {
      logger.error(errorMsg);
      logger.operation('导入失败：' + errorMsg);
    } else {
      console.error(errorMsg);
    }
    updateImportStatus(errorMsg, 'error');
    return;
  }
  
  try {
    updateImportStatus('正在导入数据...', 'info');
    
    const headerRowSelect = document.getElementById('headerRowSelect');
    if (!headerRowSelect) {
      throw new Error('无法找到字段行选择器');
    }
    
    const headerRowIndex = parseInt(headerRowSelect.value) - 1;
    console.log('字段行索引:', headerRowIndex, '数据总行数:', currentFileData.length);
    
    if (headerRowIndex >= currentFileData.length || headerRowIndex < 0) {
      throw new Error('指定的字段行超出文件范围');
    }
    
    // 创建处理后的数据
    const processedData = [
      currentFileData[headerRowIndex], // 字段行
      ...currentFileData.slice(headerRowIndex + 1) // 数据行
    ].filter(row => row && row.some(cell => cell !== null && cell !== undefined && cell.toString().trim()));
    
    console.log('处理后的数据行数:', processedData.length);
    
    if (processedData.length < 2) {
      throw new Error('没有找到有效的数据行');
    }
    
    // 验证数据完整性
    const headerLength = processedData[0].length;
    const invalidRows = processedData.slice(1).filter(row => row.length !== headerLength);
    
    if (invalidRows.length > 0) {
      const warningMsg = `发现 ${invalidRows.length} 行数据列数不匹配，将自动处理`;
      console.warn(warningMsg);
    }
    
    // 标准化数据行长度
    for (let i = 1; i < processedData.length; i++) {
      // 确保row是一个数组
      if (!Array.isArray(processedData[i])) {
        processedData[i] = [processedData[i]];
      }
      
      while (processedData[i].length < headerLength) {
        processedData[i].push('');
      }
      processedData[i] = processedData[i].slice(0, headerLength);
    }
    
    const successMsg = `导入成功：${processedData.length - 1} 条记录`;
    updateImportStatus(successMsg, 'success');
    
    // 保存文件名，避免在关闭模态框后丢失
    const fileName = currentImportFile ? currentImportFile.name : '未知文件';
    
    // 使用setTimeout确保DOM更新完成后再处理数据
    setTimeout(() => {
      try {
        // 直接调用全局的 processData 函数
        if (typeof window.processData === 'function') {
          window.processData(processedData, fileName);
          if (typeof window.updateStatus === 'function') {
            window.updateStatus(successMsg, 'success');
          }
          console.log(successMsg);
        } else {
          const errorMsg = 'window.processData函数未定义';
          console.error(errorMsg);
          
          // 尝试通过appState访问
          try {
            processData(processedData, fileName);
            updateStatus(successMsg, 'success');
            console.log(successMsg);
          } catch (e) {
            const errorMsg = '调用processData失败: ' + (e.message || e);
            console.error(errorMsg);
            updateImportStatus('数据处理异常，请刷新页面重试', 'error');
          }
        }
        
        // 数据处理完成后关闭模态框
        closeImportConfigModal();
        
        // 确保功能按钮被启用
        if (typeof window.enableDataDependentFeatures === 'function') {
          window.enableDataDependentFeatures();
        }
      } catch (err) {
        const errorMsg = '数据处理过程中发生错误: ' + (err.message || err);
        console.error(errorMsg);
        console.error('导入失败：' + errorMsg);
        updateImportStatus('数据处理异常，请刷新页面重试', 'error');
      }
    }, 100);
    
  } catch (error) {
    const errorMsg = '导入失败: ' + error.message;
    updateImportStatus(errorMsg, 'error');
    
    console.error('Import error:', error);
    
    // 即使出错也要关闭模态框
    setTimeout(() => {
      closeImportConfigModal();
    }, 50);
  }
}

// 拖拽相关
function handleDragOver(event) {
  console.log('拖拽悬停事件');
  event.preventDefault();
  const dropZone = event.currentTarget;
  if (dropZone) {
    dropZone.classList.add('active');
  }
}

function handleDragLeave(event) {
  console.log('拖拽离开事件');
  const dropZone = event.currentTarget;
  if (dropZone) {
    dropZone.classList.remove('active');
  }
}

async function handleDrop(event) {
  console.log('拖拽放置事件');
  event.preventDefault();
  const dropZone = event.currentTarget;
  if (dropZone) {
    dropZone.classList.remove('active');
  }
  
  try {
    const files = event.dataTransfer.files;
    console.log('拖拽的文件:', files);
    
    if (files.length > 0) {
      const file = files[0];
      console.log(`拖拽导入文件：${file.name} (${(file.size / 1024).toFixed(1)} KB)`);
      currentImportFile = file;
      await showImportConfigModal(file);
    } else {
      console.warn('拖拽操作未包含文件');
    }
  } catch (error) {
    const errorMsg = '拖拽文件处理失败: ' + (error.message || error);
    console.error(errorMsg);
    
    if (typeof updateStatus === 'function') {
      updateStatus(errorMsg, 'error');
    }
  }
}

// 更新表头行选择器
function updateHeaderRowSelect() {
  try {
    console.log('更新表头行选择器');
    
    const headerRowSelect = document.getElementById('headerRowSelect');
    if (!headerRowSelect) {
      console.warn('未找到表头行选择器元素');
      return;
    }
    
    // 清空现有选项
    headerRowSelect.innerHTML = '';
    
    if (!currentFileData || currentFileData.length === 0) {
      console.warn('没有可用的文件数据来更新表头行选择器');
      return;
    }
    
    // 添加表头行选项（最多显示前10行）
    const maxRows = Math.min(currentFileData.length, 10);
    for (let i = 0; i < maxRows; i++) {
      const option = document.createElement('option');
      option.value = i + 1;
      
      // 显示行号和该行的前几个字段作为预览
      const previewCells = currentFileData[i].slice(0, 3).map(cell => 
        cell ? cell.toString().substring(0, 10) : ''
      ).join(', ');
      
      option.textContent = `第${i + 1}行: ${previewCells}${currentFileData[i].length > 3 ? '...' : ''}`;
      
      // 默认选择第一行作为表头
      if (i === 0) {
        option.selected = true;
      }
      
      headerRowSelect.appendChild(option);
    }
    
    if (typeof logger !== 'undefined' && logger.log) {
      logger.log(`表头行选择器已更新，包含${maxRows}个选项`);
      logger.operation(`检测到${maxRows}行数据，可选择表头行`);
    } else {
      console.log(`表头行选择器已更新，包含${maxRows}个选项`);
    }
    
  } catch (error) {
    const errorMsg = '更新表头行选择器失败: ' + (error.message || error);
    console.error(errorMsg);
  }
}

// 确保所有需要在全局作用域中访问的函数都正确暴露
window.handleFileUpload = handleFileUpload;
window.showImportConfigModal = showImportConfigModal;
window.closeImportConfigModal = closeImportConfigModal;
window.handleWorksheetChange = handleWorksheetChange;
window.handleHeaderRowChange = handleHeaderRowChange;
window.confirmImport = confirmImport;
window.handleDragOver = handleDragOver;
window.handleDragLeave = handleDragLeave;
window.handleDrop = handleDrop;
window.updateHeaderRowSelect = updateHeaderRowSelect;
