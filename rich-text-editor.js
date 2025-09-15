/**
 * 简化的文本编辑器功能模块
 */

// 编辑器初始化
function initRichTextEditor() {
  const editor = document.getElementById('templateEditor');
  
  // 绑定编辑器事件
  editor.addEventListener('input', handleTextChange);
  editor.addEventListener('keydown', handleEditorKeydown);
  editor.addEventListener('paste', handleEditorPaste);
  
  // 绑定工具栏事件
  bindSimpleToolbarEvents();
  
  console.log('简化文本编辑器初始化完成');
}

// 绑定工具栏事件
function bindRichTextToolbarEvents() {
  // 工具栏按钮已删除，保留函数以维持兼容性
  console.log('工具栏事件绑定完成');
}

function bindSimpleToolbarEvents() {
  // 调用主工具栏绑定函数以保持一致性
  bindRichTextToolbarEvents();
}

// 处理文本内容变化
function handleTextChange() {
  const editor = document.getElementById('templateEditor');
  if (typeof appState !== 'undefined') {
    appState.template = editor.value;
  }
  
  if (typeof updatePreview === 'function') {
    updatePreview();
  }
}

// 更新字段高亮显示
function updateFieldHighlight() {
  const editor = document.getElementById('templateEditor');
  const overlay = document.getElementById('fieldHighlightOverlay');
  
  if (!editor || !overlay) return;
  
  const text = editor.value;
  
  // 创建高亮版本的文本
  const highlightedText = text.replace(/\[\[([^\]]+)\]\]/g, '<span class="field-tag">[[$1]]</span>');
  
  // 更新覆盖层内容
  overlay.innerHTML = highlightedText.replace(/\n/g, '<br>');
  
  // 同步滚动
  overlay.scrollTop = editor.scrollTop;
  overlay.scrollLeft = editor.scrollLeft;
}

// 处理键盘事件
function handleEditorKeydown(e) {
  const editor = e.target;
  const cursorPos = editor.selectionStart;
  const text = editor.value;
  
  // 支持 Tab 键缩进
  if (e.key === 'Tab') {
    e.preventDefault();
    const start = editor.selectionStart;
    const end = editor.selectionEnd;
    
    // 插入制表符
    editor.value = editor.value.substring(0, start) + '  ' + editor.value.substring(end);
    editor.selectionStart = editor.selectionEnd = start + 2;
    
    handleTextChange();
    return;
  }
  
  // 字段保护机制：防止部分删除字段变量
  if (e.key === 'Backspace' || e.key === 'Delete') {
    const result = protectFieldVariables(editor, e.key, cursorPos);
    if (result.preventDefault) {
      e.preventDefault();
      if (result.newValue !== undefined) {
        editor.value = result.newValue;
        editor.selectionStart = editor.selectionEnd = result.newCursor;
        handleTextChange();
      }
    }
  }
}


// 字段保护机制：防止部分删除字段变量
function protectFieldVariables(editor, key, cursorPos) {
  const text = editor.value;
  
  // 查找所有字段变量
  const fieldRegex = /\[\[([^\]]+)\]\]/g;
  let match;
  const fields = [];
  
  while ((match = fieldRegex.exec(text)) !== null) {
    fields.push({
      start: match.index,
      end: match.index + match[0].length,
      content: match[0]
    });
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

// 处理粘贴事件
function handleEditorPaste(e) {
  // 只允许粘贴纯文本
  e.preventDefault();
  
  const clipboardData = e.clipboardData || window.clipboardData;
  const pastedData = clipboardData.getData('text/plain');
  
  const editor = e.target;
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  
  // 插入粘贴的文本
  editor.value = editor.value.substring(0, start) + pastedData + editor.value.substring(end);
  editor.selectionStart = editor.selectionEnd = start + pastedData.length;
  
  handleTextChange();
}

// 插入字段变量
function insertFieldVariable(fieldName) {
  const editor = document.getElementById('templateEditor');
  const fieldVariable = `[[${fieldName}]]`;
  
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  
  // 在光标位置插入字段变量
  editor.value = editor.value.substring(0, start) + fieldVariable + editor.value.substring(end);
  
  // 设置光标位置到插入内容之后
  editor.selectionStart = editor.selectionEnd = start + fieldVariable.length;
  
  // 焦点回到编辑器
  editor.focus();
  
  handleTextChange();
}

// 将函数暴露到全局作用域
window.insertFieldVariable = insertFieldVariable;

// 清空编辑器内容
function clearEditor() {
  if (confirm('确定要清空编辑器内容吗？')) {
    const editor = document.getElementById('templateEditor');
    editor.value = '';
    editor.focus();
    handleTextChange();
  }
}

// 兼容性函数（保持与富文本版本的接口一致）
function handleRichTextChange() {
  handleTextChange();
}

// 虚拟函数，保持兼容性
function insertHtmlAtCursor(html) {
  // 对于纯文本编辑器，清理HTML标签
  const textContent = html.replace(/<[^>]*>/g, '');
  insertTextAtCursor(textContent);
}

function insertTextAtCursor(text) {
  const editor = document.getElementById('templateEditor');
  const start = editor.selectionStart;
  const end = editor.selectionEnd;
  
  editor.value = editor.value.substring(0, start) + text + editor.value.substring(end);
  editor.selectionStart = editor.selectionEnd = start + text.length;
  editor.focus();
  
  handleTextChange();
}