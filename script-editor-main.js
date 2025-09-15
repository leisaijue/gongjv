/**
 * 话术编辑器主应用逻辑模块
 */



  


// 应用状态
const appState = { 
  data: [], 
  fields: [], 
  currentCustomer: null, 
  template: '', 
  filteredCustomers: [],
  selectedCustomerIndex: -1,
  isDropdownOpen: false,
  nameColumn: '',
  calcFields: [], // 计算字段
  scriptHistory: [] // 历史话术记录
};



// DOM 选择器简化函数
const $ = id => {
  const element = document.getElementById(id);
  if (!element) {
    console.warn(`未找到ID为 "${id}" 的元素`);
  }
  return element;
};

// 应用初始化
function initApp() {
  console.log('初始化应用...');
  
  try {
    bindEvents();
    console.log('事件绑定完成');
    
    initRichTextEditor();
    console.log('富文本编辑器初始化完成');
    
    loadHistoryFromStorage();
    console.log('历史记录加载完成');
    
    // 添加全局错误处理
    window.addEventListener('error', function(e) {
      console.error('全局错误捕获:', e.error);
    });
    
    console.log('话术编辑器初始化完成');
  } catch (error) {
    console.error('应用初始化失败:', error);
  }
}

// 绑定所有事件监听器
function bindEvents() {
  console.log('绑定事件监听器...');
  
  // 文件上传功能
  const uploadBtn = $('uploadBtn');
  const fileInput = $('fileInput');
  
  if (uploadBtn && fileInput) {
    uploadBtn.addEventListener('click', () => {
      fileInput.click();
    });
    
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        handleFileUpload(e);
      }
    });
  }
  
  // 编辑器相关
  const templateEditor = $('templateEditor');
  if (templateEditor) {
    templateEditor.addEventListener('input', () => {
      handleTextChange();
    });
  }
  
  // 客户搜索相关
  const customerSearchInput = $('customerSearchInput');
  if (customerSearchInput) {
    customerSearchInput.addEventListener('input', (e) => {
      handleCustomerSearch(e);
    });
    customerSearchInput.addEventListener('focus', handleSearchFocus);
    customerSearchInput.addEventListener('click', handleSearchClick);
    customerSearchInput.addEventListener('blur', handleSearchBlur);
    customerSearchInput.addEventListener('keydown', handleSearchKeydown);
  }
  
  // 操作按钮
  const copyCurrentBtn = $('copyCurrentBtn');
  const batchExportBtn = $('batchExportBtn');
  const clearDataBtn = $('clearDataBtn');
  const loadSampleBtn = $('loadSampleBtn');
  const nameColumnSelect = $('nameColumnSelect');
  
  if (copyCurrentBtn) copyCurrentBtn.addEventListener('click', () => {
    copyCurrentScript();
  });
  if (batchExportBtn) batchExportBtn.addEventListener('click', () => {
    batchExport();
  });
  if (clearDataBtn) clearDataBtn.addEventListener('click', () => {
    clearAllData();
  });
  if (loadSampleBtn) loadSampleBtn.addEventListener('click', () => {
    loadSampleData();
  });
  if (nameColumnSelect) nameColumnSelect.addEventListener('change', (e) => {
    handleNameColumnChange(e);
  });
  
  // 历史话术记录相关事件
  const historyBtn = $('historyBtn');
  const clearHistoryBtn = $('clearHistoryBtn');
  const historySearchInput = $('historySearchInput');
  
  if (historyBtn) historyBtn.addEventListener('click', () => {
    showHistoryModal();
  });
  if (clearHistoryBtn) clearHistoryBtn.addEventListener('click', () => {
    clearAllHistory();
  });
  if (historySearchInput) historySearchInput.addEventListener('input', (e) => {
    filterHistoryList(e);
    logger.operation('搜索历史话术：' + e.target.value);
  });
  
  // 字段计算相关事件
  const addCalcFieldBtn = $('addCalcFieldBtn');
  const addCalcFieldConfirmBtn = $('addCalcFieldConfirmBtn');
  
  if (addCalcFieldBtn) addCalcFieldBtn.addEventListener('click', () => {
    showCalcFieldModal();
  });
  if (addCalcFieldConfirmBtn) addCalcFieldConfirmBtn.addEventListener('click', () => {
    addCalculatedField();
  });
  
  // 富文本编辑器工具栏事件
  bindRichTextToolbarEvents();
  
  // 导入配置相关事件
  const worksheetSelect = $('worksheetSelect');
  const headerRowSelect = $('headerRowSelect');
  
  if (worksheetSelect) worksheetSelect.addEventListener('change', (e) => {
    handleWorksheetChange(e);
  });
  if (headerRowSelect) headerRowSelect.addEventListener('change', (e) => {
    handleHeaderRowChange(e);
  });
  
  // 确认导入按钮事件绑定
  const confirmImportBtn = $('confirmImportBtn');
  if (confirmImportBtn) {
    // 移除所有可能已存在的事件监听器
    const newConfirmImportBtn = confirmImportBtn.cloneNode(true);
    confirmImportBtn.parentNode.replaceChild(newConfirmImportBtn, confirmImportBtn);
    
    // 重新获取引用并绑定事件
    const finalConfirmImportBtn = $('confirmImportBtn');
    finalConfirmImportBtn.addEventListener('click', function(e) {
      console.log('确认导入按钮被点击');
      e.preventDefault();
      e.stopPropagation();
      
      try {
        // 调用confirmImport函数
        if (typeof confirmImport === 'function') {
          confirmImport();
        } else if (typeof window.confirmImport === 'function') {
          window.confirmImport();
        } else {
          throw new Error('confirmImport函数未定义');
        }
      } catch (err) {
        console.error('导入功能异常：' + err.message);
        updateStatus('导入功能异常，请刷新页面重试', 'error');
      }
    });
  } else {
    console.error('未找到确认导入按钮元素');
  }
  
  // 字段搜索事件
  const fieldSearchInput = $('fieldSearchInput');
  if (fieldSearchInput) {
    fieldSearchInput.addEventListener('input', (e) => {
      handleFieldSearch(e);
    });
    fieldSearchInput.addEventListener('keyup', handleFieldSearch);
  }
  
  // 全局点击事件
  document.addEventListener('click', handleDocumentClick);
  
  // 文件拖拽事件
  const fileDropZone = $('fileDropZone');
  if (fileDropZone) {
    fileDropZone.addEventListener('dragover', (e) => {
      handleDragOver(e);
      logger.operation('文件拖拽悬停');
    });
    fileDropZone.addEventListener('dragleave', (e) => {
      handleDragLeave(e);
      logger.operation('文件拖拽离开');
    });
    fileDropZone.addEventListener('drop', (e) => {
      handleDrop(e);
      const fileName = e.dataTransfer.files[0]?.name || '未知文件';
      logger.operation('拖拽放置文件：' + fileName);
    });
    fileDropZone.addEventListener('click', () => {
      logger.operation('点击文件拖拽区域（功能已禁用）');
      alert('文件上传功能已被禁用');
    });
  }
  
  // 添加全局错误捕获
  window.addEventListener('error', (event) => {
    console.error('全局错误：' + event.message + ' 在 ' + event.filename + ':' + event.lineno);
    return false;
  });
  
  // 添加Promise错误捕获
  window.addEventListener('unhandledrejection', (event) => {
    console.error('未处理的Promise错误：' + (event.reason?.message || event.reason || '未知错误'));
    return false;
  });
  
  console.log('事件绑定完成');
}

// 数据处理主函数
window.processData = function processData(data, fileName) {
  console.log('开始处理数据:', data, '文件名:', fileName);
  if (!data || data.length < 2) {
    updateStatus('文件为空或格式错误', 'error');
    return;
  }
  
  const headers = data[0];
  const rows = data.slice(1);
  
  console.log('表头:', headers);
  console.log('数据行数:', rows.length);
  
  // 验证数据完整性
  if (!headers || headers.length === 0) {
    updateStatus('未找到有效的字段名', 'error');
    return;
  }
  
  // 转换为对象格式
  const processedData = rows.map((row, index) => {
    // 添加调试日志
    if (index < 3) {
      console.log(`处理第${index}行数据:`, row);
    }
    
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] !== undefined ? row[index] : '';
    });
    return obj;
  }).filter((row, index) => {
    const keep = Object.values(row).some(value => value !== null && value !== undefined && value.toString().trim());
    if (index < 3) {
      console.log(`过滤第${index}行:`, row, '保留:', keep);
    }
    return keep;
  });
  
  console.log('处理后的数据:', processedData);
  
  if (processedData.length === 0) {
    updateStatus('没有找到有效的数据行', 'error');
    return;
  }
  
  // 更新应用状态
  appState.data = processedData;
  appState.fields = headers.map((name, index) => ({ name, index }));
  appState.currentCustomer = null;
  appState.filteredCustomers = [...processedData];
  appState.calcFields = []; // 清空计算字段
  
  // 应用计算字段（如果有的话）
  if (appState.calcFields.length > 0) {
    applyCalculatedFields();
  }
  
  // 更新界面
  updateDataOverview(fileName, processedData.length, headers.length);
  updateFieldsList();
  updateNameColumnSelector();
  updateCustomerOptions();
  updatePreview();
  enableDataDependentFeatures();
  
  updateStatus(`成功加载 ${processedData.length} 条记录，${headers.length} 个字段`, 'success');
}

// 更新数据概览
function updateDataOverview(fileName, recordCount, fieldCount) {
  console.log('更新数据概览:', fileName, recordCount, fieldCount);
  const fileNameElement = $('fileName');
  const recordCountElement = $('recordCount');
  const fieldCountElement = $('fieldCount');
  
  if (fileNameElement) fileNameElement.textContent = fileName || '未命名文件';
  if (recordCountElement) recordCountElement.textContent = recordCount || 0;
  if (fieldCountElement) fieldCountElement.textContent = fieldCount || 0;
}

// 更新字段列表
function updateFieldsList() {
  console.log('更新字段列表');
  const fieldsList = $('fieldsList');
  if (!fieldsList) {
    console.error('未找到字段列表容器');
    return;
  }
  
  if (appState.fields.length === 0 && appState.calcFields.length === 0) {
    fieldsList.innerHTML = `
      <div class="file-drop-zone" id="fileDropZone">
        <div style="font-size: 24px; margin-bottom: 8px;">📁</div>
        <div style="font-weight: 500; margin-bottom: 4px;">拖拽文件到此处</div>
        <div style="font-size: 12px;">支持 CSV, TSV, XLSX 格式</div>
      </div>
    `;
    
    // 重新绑定拖拽事件
    const dropZone = $('fileDropZone');
    if (dropZone) {
      dropZone.addEventListener('dragover', handleDragOver);
      dropZone.addEventListener('dragleave', handleDragLeave);
      dropZone.addEventListener('drop', handleDrop);
      dropZone.addEventListener('click', () => {
        const fileInput = $('fileInput');
        if (fileInput) fileInput.click();
      });
    }
    return;
  }
  
  fieldsList.innerHTML = '';
  
  // 添加原始字段
  appState.fields.forEach(field => {
    const fieldItem = document.createElement('div');
    fieldItem.className = 'field-item';
    fieldItem.innerHTML = `
      <div class="field-name">${field.name}</div>
      <button class="field-insert" onclick="insertFieldInEditor('${field.name}')">+</button>
    `;
    fieldsList.appendChild(fieldItem);
  });
  
  // 添加计算字段
  appState.calcFields.forEach(calcField => {
    const fieldItem = document.createElement('div');
    fieldItem.className = 'field-calc-item';
    fieldItem.innerHTML = `
      <div>
        <div class="calc-field-name">🧠 ${calcField.name}</div>
        <div class="calc-field-formula">${calcField.formula}</div>
      </div>
      <button class="field-insert" onclick="insertFieldInEditor('${calcField.name}')">+</button>
    `;
    fieldsList.appendChild(fieldItem);
  });
}

// 在编辑器中插入字段
function insertFieldInEditor(fieldName) {
  console.log('插入字段到编辑器:', fieldName);
  console.log('insertFieldVariable函数类型:', typeof insertFieldVariable);
  
  // 检查编辑器元素是否存在
  const editor = document.getElementById('templateEditor');
  console.log('编辑器元素:', editor);
  
  if (typeof insertFieldVariable === 'function') {
    console.log('调用insertFieldVariable函数');
    insertFieldVariable(fieldName);
    console.log('insertFieldVariable函数调用完成');
  } else {
    console.error('insertFieldVariable函数未定义');
    // 如果函数未定义，直接在这里实现插入逻辑
    if (editor) {
      const fieldVariable = `[[${fieldName}]]`;
      const start = editor.selectionStart;
      const end = editor.selectionEnd;
      
      editor.value = editor.value.substring(0, start) + fieldVariable + editor.value.substring(end);
      editor.selectionStart = editor.selectionEnd = start + fieldVariable.length;
      editor.focus();
      
      console.log('直接插入字段变量:', fieldVariable);
      handleTextChange();
    }
  }
}

// 处理文本内容变化
function handleTextChange() {
  console.log('文本内容变化');
  const editor = $('templateEditor');
  if (editor) {
    appState.template = editor.value;
    updatePreview();
  }
}

// 更新客户名称列选择器
function updateNameColumnSelector() {
  console.log('更新客户名称列选择器');
  const nameColumnSelector = $('nameColumnSelector');
  const nameColumnSelect = $('nameColumnSelect');
  
  if (!nameColumnSelect) {
    console.log('客户名称列选择器不存在，跳过更新');
    return;
  }
  
  if (appState.fields.length === 0) {
    if (nameColumnSelector) nameColumnSelector.style.display = 'none';
    return;
  }
  
  if (nameColumnSelector) nameColumnSelector.style.display = 'block';
  nameColumnSelect.innerHTML = '<option value="">请选择客户名称列...</option>';
  
  // 添加原始字段
  appState.fields.forEach(field => {
    const option = document.createElement('option');
    option.value = field.name;
    option.textContent = field.name;
    nameColumnSelect.appendChild(option);
  });
  
  // 添加计算字段
  appState.calcFields.forEach(calcField => {
    const option = document.createElement('option');
    option.value = calcField.name;
    option.textContent = `🧮 ${calcField.name}`;
    nameColumnSelect.appendChild(option);
  });
  
  // 智能选择默认名称列
  const nameFields = ['客户名称', '姓名', '名称', 'name', 'customer', 'client', '客户'];
  const autoField = appState.fields.find(field => 
    nameFields.some(name => field.name.toLowerCase().includes(name.toLowerCase()))
  );
  
  if (autoField) {
    nameColumnSelect.value = autoField.name;
    appState.nameColumn = autoField.name;
    updateCustomerOptions();
  }
}

// 客户搜索相关函数
function handleCustomerSearch() {
  console.log('处理客户搜索');
  const customerSearchInput = $('customerSearchInput');
  if (!customerSearchInput) return;
  
  const searchTerm = customerSearchInput.value.toLowerCase().trim();
  
  if (!searchTerm) {
    appState.filteredCustomers = [...appState.data];
  } else {
    appState.filteredCustomers = appState.data.filter(customer => {
      return Object.values(customer).some(value => 
        value && value.toString().toLowerCase().includes(searchTerm)
      );
    });
  }
  
  appState.selectedCustomerIndex = -1;
  updateCustomerOptions();
  showCustomerDropdown();
}

function handleSearchFocus() {
  console.log('搜索框获得焦点');
  showCustomerDropdown();
}

function handleSearchClick() {
  console.log('搜索框被点击');
  showCustomerDropdown();
}

function handleSearchBlur() {
  console.log('搜索框失去焦点');
  // 延迟隐藏，以便处理点击选项的情况
  setTimeout(() => {
    hideCustomerDropdown();
  }, 200);
}

function handleSearchKeydown(e) {
  console.log('搜索框按键:', e.key);
  const options = $('customerOptions');
  if (!options) return;
  
  const optionElements = options.querySelectorAll('.customer-option');
  
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    appState.selectedCustomerIndex = Math.min(
      appState.selectedCustomerIndex + 1, 
      appState.filteredCustomers.length - 1
    );
    updateHighlight();
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    appState.selectedCustomerIndex = Math.max(appState.selectedCustomerIndex - 1, -1);
    updateHighlight();
  } else if (e.key === 'Enter') {
    e.preventDefault();
    if (appState.selectedCustomerIndex >= 0) {
      selectCustomer(appState.filteredCustomers[appState.selectedCustomerIndex]);
    }
  } else if (e.key === 'Escape') {
    hideCustomerDropdown();
    const customerSearchInput = $('customerSearchInput');
    if (customerSearchInput) customerSearchInput.blur();
  }
}

function showCustomerDropdown() {
  console.log('显示客户下拉列表');
  if (appState.data.length === 0) return;
  
  appState.isDropdownOpen = true;
  updateCustomerOptions();
  
  const customerOptions = $('customerOptions');
  if (customerOptions) {
    customerOptions.classList.add('show');
  }
}

function hideCustomerDropdown() {
  console.log('隐藏客户下拉列表');
  appState.isDropdownOpen = false;
  
  const customerOptions = $('customerOptions');
  if (customerOptions) {
    customerOptions.classList.remove('show');
  }
}

function updateCustomerOptions() {
  console.log('更新客户选项');
  const options = $('customerOptions');
  if (!options) {
    console.error('未找到客户选项容器');
    return;
  }
  
  if (appState.filteredCustomers.length === 0) {
    options.innerHTML = '<div class="no-results">没有找到匹配的客户</div>';
    return;
  }
  
  options.innerHTML = '';
  
  appState.filteredCustomers.slice(0, 10).forEach((customer, index) => {
    const option = document.createElement('div');
    option.className = 'customer-option';
    
    const displayName = getCustomerDisplayName(customer);
    option.textContent = displayName;
    
    option.addEventListener('click', () => selectCustomer(customer));
    options.appendChild(option);
  });
  
  if (appState.filteredCustomers.length > 10) {
    const moreOption = document.createElement('div');
    moreOption.className = 'no-results';
    moreOption.textContent = `还有 ${appState.filteredCustomers.length - 10} 个结果...`;
    options.appendChild(moreOption);
  }
  
  updateHighlight();
}

function updateHighlight() {
  console.log('更新高亮选项');
  const options = $('customerOptions');
  if (!options) return;
  
  const optionElements = options.querySelectorAll('.customer-option');
  
  optionElements.forEach((element, index) => {
    if (index === appState.selectedCustomerIndex) {
      element.classList.add('highlighted');
    } else {
      element.classList.remove('highlighted');
    }
  });
}

function selectCustomer(customer) {
  console.log('选择客户:', customer);
  appState.currentCustomer = customer;
  
  const displayName = getCustomerDisplayName(customer);
  const customerSearchInput = $('customerSearchInput');
  if (customerSearchInput) {
    customerSearchInput.value = displayName;
  }
  
  hideCustomerDropdown();
  updatePreview();
  
  // 标记客户选项为已选择
  const options = $('customerOptions');
  if (options) {
    const optionElements = options.querySelectorAll('.customer-option');
    optionElements.forEach(element => {
      if (element.textContent === displayName) {
        element.classList.add('selected');
      } else {
        element.classList.remove('selected');
      }
    });
  }
  
  enableGenerationFeatures();
  updateStatus(`已选择客户：${displayName}`, 'success');
}

function getCustomerDisplayName(customer) {
  console.log('获取客户显示名称:', customer);
  if (appState.nameColumn && customer[appState.nameColumn]) {
    return customer[appState.nameColumn];
  }
  
  // 尝试常见的名称字段
  const nameFields = ['客户名称', '姓名', '名称', 'name', 'customer', 'client'];
  for (const field of nameFields) {
    if (customer[field]) {
      return customer[field];
    }
  }
  
  // 返回第一个非空字段
  for (const [key, value] of Object.entries(customer)) {
    if (value && value.toString().trim()) {
      return `${key}: ${value}`;
    }
  }
  
  return '未知客户';
}

// 客户名称列变化处理
function handleNameColumnChange() {
  console.log('客户名称列变化');
  const nameColumnSelect = $('nameColumnSelect');
  if (!nameColumnSelect) return;
  
  const selectedColumn = nameColumnSelect.value;
  appState.nameColumn = selectedColumn;
  
  updateCustomerOptions();
  
  if (appState.currentCustomer) {
    const displayName = getCustomerDisplayName(appState.currentCustomer);
    const customerSearchInput = $('customerSearchInput');
    if (customerSearchInput) {
      customerSearchInput.value = displayName;
    }
  }
  
  updateStatus(`客户名称列已设置为：${selectedColumn}`, 'info');
}

// 字段搜索功能
function handleFieldSearch(e) {
  console.log('字段搜索:', e.target.value);
  const searchTerm = e.target.value.toLowerCase().trim();
  const fieldsList = $('fieldsList');
  if (!fieldsList) return;
  
  const fieldItems = fieldsList.querySelectorAll('.field-item, .field-calc-item');
  
  fieldItems.forEach(item => {
    const fieldName = item.querySelector('.field-name, .calc-field-name');
    if (fieldName) {
      const fieldNameText = fieldName.textContent.toLowerCase();
      if (!searchTerm || fieldNameText.includes(searchTerm)) {
        item.style.display = '';
      } else {
        item.style.display = 'none';
      }
    }
  });
}

// 全局点击事件处理
function handleDocumentClick(e) {
  console.log('文档点击事件');
  const customerOptions = $('customerOptions');
  const customerSearchInput = $('customerSearchInput');
  
  if (customerOptions && customerSearchInput && 
      !customerOptions.contains(e.target) && e.target !== customerSearchInput) {
    hideCustomerDropdown();
  }
}

// 启用依赖数据的功能
window.enableDataDependentFeatures = function enableDataDependentFeatures() {
  console.log('启用依赖数据的功能');
  const clearDataBtn = $('clearDataBtn');
  const customerSearchInput = $('customerSearchInput');
  const addCalcFieldBtn = $('addCalcFieldBtn');
  const fieldSearchContainer = $('fieldSearchContainer');
  
  if (clearDataBtn) clearDataBtn.disabled = false;
  if (customerSearchInput) {
    customerSearchInput.disabled = false;
    customerSearchInput.placeholder = '搜索并选择客户...';
  }
  if (addCalcFieldBtn) addCalcFieldBtn.disabled = false;
  if (fieldSearchContainer) fieldSearchContainer.style.display = 'block';
}

// 启用生成相关功能
function enableGenerationFeatures() {
  console.log('启用生成相关功能');
  const copyCurrentBtn = $('copyCurrentBtn');
  const batchExportBtn = $('batchExportBtn');
  
  if (copyCurrentBtn) copyCurrentBtn.disabled = false;
  if (batchExportBtn) batchExportBtn.disabled = false;
}

// 清空所有数据
function clearAllData() {
  console.log('清空所有数据');
  if (confirm('确定要清空所有数据吗？此操作将清除已上传的数据和计算字段。')) {
    appState.data = [];
    appState.fields = [];
    appState.currentCustomer = null;
    appState.filteredCustomers = [];
    appState.nameColumn = '';
    appState.calcFields = [];
    
    updateDataOverview('未上传', 0, 0);
    updateFieldsList();
    updateNameColumnSelector();
    updateCustomerOptions();
    updatePreview();
    
    // 禁用功能
    const clearDataBtn = $('clearDataBtn');
    const copyCurrentBtn = $('copyCurrentBtn');
    const batchExportBtn = $('batchExportBtn');
    const customerSearchInput = $('customerSearchInput');
    const addCalcFieldBtn = $('addCalcFieldBtn');
    const fieldSearchContainer = $('fieldSearchContainer');
    
    if (clearDataBtn) clearDataBtn.disabled = true;
    if (copyCurrentBtn) copyCurrentBtn.disabled = true;
    if (batchExportBtn) batchExportBtn.disabled = true;
    if (customerSearchInput) {
      customerSearchInput.disabled = true;
      customerSearchInput.placeholder = '请先上传数据...';
      customerSearchInput.value = '';
    }
    if (addCalcFieldBtn) addCalcFieldBtn.disabled = true;
    if (fieldSearchContainer) fieldSearchContainer.style.display = 'none';
    
    updateStatus('数据已清空', 'success');
  }
}

// 复制当前话术
async function copyCurrentScript() {
  console.log('复制当前话术');
  if (!appState.currentCustomer || !appState.template) {
    alert('请先选择客户并编辑话术模板');
    return;
  }
  
  const renderedScript = renderTemplate(appState.template, appState.currentCustomer);
  
  try {
    await navigator.clipboard.writeText(renderedScript);
    
    // 添加到历史记录
    addToHistory(appState.currentCustomer, appState.template, renderedScript);
    
    updateStatus('话术已复制到剪贴板', 'success');
  } catch (error) {
    // 降级到传统方法
    const textArea = document.createElement('textarea');
    textArea.value = renderedScript;
    document.body.appendChild(textArea);
    textArea.select();
    document.execCommand('copy');
    document.body.removeChild(textArea);
    
    addToHistory(appState.currentCustomer, appState.template, renderedScript);
    updateStatus('话术已复制到剪贴板', 'success');
  }
}

// 批量导出
function batchExport() {
  console.log('批量导出');
  if (appState.data.length === 0 || !appState.template) {
    alert('请先上传数据并编辑话术模板');
    return;
  }
  
  const results = appState.data.map(customer => {
    const displayName = getCustomerDisplayName(customer);
    const renderedScript = renderTemplate(appState.template, customer);
    
    return {
      '客户名称': displayName,
      '话术内容': renderedScript,
      '生成时间': new Date().toLocaleString('zh-CN'),
      ...customer // 包含原始数据
    };
  });
  
  downloadAsExcel(results, '批量话术导出');
  updateStatus(`已导出 ${results.length} 条话术`, 'success');
}

// 状态更新函数
window.updateStatus = function updateStatus(message, type = 'info') {
  console.log('更新状态:', message, '类型:', type);
  const statusText = $('statusText');
  const templateStatus = $('templateStatus');
  
  if (statusText) statusText.textContent = message;
  if (templateStatus) templateStatus.className = `status ${type}`;
  
  // 3秒后恢复默认状态
  setTimeout(() => {
    if (appState.data.length === 0) {
      if (statusText) statusText.textContent = '等待上传数据';
      if (templateStatus) templateStatus.className = 'status info';
    } else if (!appState.currentCustomer) {
      if (statusText) statusText.textContent = '请选择客户';
      if (templateStatus) templateStatus.className = 'status warning';
    } else {
      if (statusText) statusText.textContent = '就绪';
      if (templateStatus) templateStatus.className = 'status success';
    }
  }, 3000);
}

// 加载示例数据
async function loadSampleData() {
  console.log('加载示例数据');
  try {
    const response = await fetch('./test-chinese.csv');
    if (!response.ok) {
      throw new Error('示例文件加载失败');
    }
    
    const csvText = await response.text();
    const lines = csvText.trim().split('\n');
    const data = lines.map(line => {
      return line.split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
    });
    
    processData(data, 'test-chinese.csv (示例)');
  } catch (error) {
    console.error('加载示例数据失败:', error);
    updateStatus('示例数据加载失败', 'error');
  }
}

// Excel 下载功能
function downloadAsExcel(data, filename) {
  console.log('下载Excel文件');
  if (typeof XLSX === 'undefined') {
    alert('Excel导出功能未加载');
    return;
  }
  
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '话术数据');
  
  XLSX.writeFile(workbook, `${filename}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// 应用初始化
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOM内容加载完成，初始化应用');
  initApp();
});