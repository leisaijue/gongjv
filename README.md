# 话术编辑器项目结构

## 📁 目录结构
```
工具箱代码/
├── css/
│   └── script-editor.css          # 样式文件
├── js/
│   ├── data-import.js             # 数据导入功能
│   ├── rich-text-editor.js        # 简化文本编辑器
│   ├── script-editor-main.js      # 主应用逻辑
│   └── template-render.js         # 模板渲染和历史记录
├── customer-card.html             # 客户信息卡生成器
├── customer-search.html           # 客户信息搜索
├── image-gallery.html             # 批量图片展示器
├── index.html                     # 首页
├── login.html                     # 登录页面
├── script-editor.html             # 话术编辑器（主要功能）
└── test-chinese.csv               # 中文测试数据
```

## 🎯 核心文件说明

### 话术编辑器相关
- **script-editor.html**: 主界面，包含三栏布局（数据面板、编辑器、预览）
- **js/script-editor-main.js**: 核心业务逻辑，数据处理和状态管理
- **js/rich-text-editor.js**: 简化的文本编辑器，支持字段变量插入
- **js/template-render.js**: 模板渲染引擎和历史记录管理
- **js/data-import.js**: 支持CSV/Excel文件导入和编码处理
- **css/script-editor.css**: 现代化UI样式和响应式布局

### 其他工具
- **customer-search.html**: 客户信息搜索工具
- **customer-card.html**: 客户信息卡片生成器
- **image-gallery.html**: 批量图片展示工具
- **index.html**: 工具集合首页
- **login.html**: 统一登录验证

## 🚀 功能特性

### 话术编辑器
- ✅ 普通文本编辑器（已简化）
- ✅ 字段变量插入 [[字段名]]
- ✅ CSV/Excel数据导入
- ✅ 实时预览渲染
- ✅ 历史话术记录
- ✅ 批量导出功能
- ✅ 计算字段支持
- ✅ 中文编码优化
- ✅ 现代化UI设计
- ✅ 响应式布局

### 已移除功能
- ❌ 富文本格式化（字体、颜色、对齐等）
- ❌ 表格插入和编辑
- ❌ 图片插入功能
- ❌ 复杂的工具栏

## 📋 最近更新
- 将富文本编辑器简化为普通文本编辑器
- 删除不必要的模块和文件
- 优化代码结构和性能
- 改进中文编码支持
- 更新UI设计为现代化风格

## 🛠️ 开发说明
- 使用原生JavaScript，无框架依赖
- 模块化代码组织
- 支持现代浏览器
- 移动端友好设计