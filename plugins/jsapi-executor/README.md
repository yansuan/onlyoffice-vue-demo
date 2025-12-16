# JsAPI Executor 插件 - 社区版解决方案

这个插件允许你从外部 Vue 应用执行 JsAPI 代码，是社区版中实现类似 OnlyOffice Playground 功能的解决方案。

## 工作原理

1. **Vue 应用**：将 JsAPI 代码和请求 ID 写入 `localStorage`
2. **插件监听**：插件定期检查 `localStorage` 中的新请求
3. **执行代码**：插件在编辑器内部执行 JsAPI 代码（可以访问 `Api` 对象）
4. **返回结果**：插件将执行结果写回 `localStorage`，Vue 应用读取并显示

## 部署步骤

1. **复制插件到 DocumentServer**
   ```bash
   # 将整个 jsapi-executor 目录复制到 DocumentServer 的插件目录
   cp -r plugins/jsapi-executor /var/www/onlyoffice/documentserver/sdkjs-plugins/
   ```

2. **重启 DocumentServer**
   ```bash
   # Docker 方式
   docker restart documentserver
   
   # 或系统服务方式
   supervisorctl restart ds:docservice
   ```

3. **在 Vue 应用中使用**
   - 打开 Vue 应用
   - 在代码编辑器中输入 JsAPI 代码
   - 点击"执行 JsAPI"按钮

## 使用示例

### 示例 1：添加段落
```javascript
var doc = Api.GetDocument();
var para = Api.CreateParagraph();
para.AddText("这是通过 JsAPI 添加的文字！");
para.SetJc("center"); // 居中对齐
doc.Push(para);
```

### 示例 2：添加批注
```javascript
var doc = Api.GetDocument();
var range = doc.GetRange();
var comment = range.AddComment("这是批注内容");
comment.SetAuthor("JsAPI User");
comment.SetDisplayName("JsAPI User");
```

### 示例 3：搜索文字
```javascript
var doc = Api.GetDocument();
var searchText = "要搜索的文字";
var paragraphs = doc.GetElements(Api.c_oAscElementTypeParagraph);

for (var i = 0; i < paragraphs.length; i++) {
  var para = paragraphs[i];
  var paraText = para.GetText();
  
  if (paraText && paraText.indexOf(searchText) !== -1) {
    Api.GetConsole().Log("找到文字在第 " + (i + 1) + " 段");
    break;
  }
}
```

## 与 Playground 的区别

| 特性 | OnlyOffice Playground | 本插件（社区版） |
|------|----------------------|-----------------|
| 需要版本 | Developer Edition | Community Edition |
| 执行方式 | `connector.callCommand()` | 插件 + localStorage |
| 代码注入 | 直接注入 | 通过插件间接执行 |
| 功能 | 完整支持 | 完整支持 |

## 注意事项

1. **插件必须已部署并启用**：确保插件已正确部署到 DocumentServer
2. **代码在编辑器内部执行**：所有代码都在编辑器环境中运行，可以访问 `Api` 对象
3. **错误处理**：代码执行错误会被捕获并返回给 Vue 应用
4. **异步执行**：执行是异步的，需要等待插件返回结果

## 技术细节

### localStorage 通信协议

**请求格式**：
```json
{
  "code": "var doc = Api.GetDocument(); ...",
  "requestId": "jsapi-1234567890",
  "timestamp": 1234567890
}
```

**响应格式**：
```json
{
  "requestId": "jsapi-1234567890",
  "success": true,
  "result": null,
  "error": null
}
```

## 故障排除

1. **插件未加载**：检查 DocumentServer 日志，确认插件已正确加载
2. **执行超时**：确保插件正在运行，检查浏览器控制台
3. **代码错误**：查看执行结果中的错误信息，检查 JsAPI 代码语法

