# OnlyOffice JsAPI 使用说明

根据 [OnlyOffice Text Document API 文档](https://api.onlyoffice.com/zh-CN/docs/office-api/usage-api/text-document-api/)，JsAPI 只能在编辑器内部运行。

## 重要说明

**JsAPI 无法从外部 Vue 应用直接调用**，因为：
1. JsAPI 运行在编辑器的 JavaScript 环境中
2. 外部应用无法直接访问编辑器的内部 API
3. 必须通过某种方式在编辑器内部执行 JsAPI 代码

## 使用方式

### 方式 1：在宏中使用 JsAPI（手动运行）

在编辑器中打开"开发工具" -> "宏"，创建宏并运行。

### 方式 2：在插件中使用 JsAPI（推荐）

插件使用 Plugins API 作为入口，内部使用 JsAPI 实现功能。这是唯一可以从外部按钮触发 JsAPI 的方式。

### 方式 3：在 Document Builder 中使用（服务器端）

在服务器端使用 Document Builder 处理文档，不涉及前端。

## JsAPI 核心接口

根据文档，主要的 JsAPI 接口包括：

- `Api.GetDocument()` - 获取文档对象
- `Api.CreateParagraph()` - 创建段落
- `Api.GetRange()` - 获取选区
- `Api.GetElement(index)` - 获取文档元素
- 等等...

详见：[Text Document API 文档](https://api.onlyoffice.com/zh-CN/docs/office-api/usage-api/text-document-api/)

