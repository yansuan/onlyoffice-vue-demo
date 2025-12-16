// OnlyOffice JsAPI 使用示例
// 根据文档：https://api.onlyoffice.com/zh-CN/docs/office-api/usage-api/text-document-api/
//
// 注意：这些代码必须在编辑器内部运行（宏或插件中）
// 无法从外部 Vue 应用直接调用

// ============================================
// 1. 获取文档对象
// ============================================
var doc = Api.GetDocument()

// ============================================
// 2. 获取或创建段落
// ============================================
// 获取第一个段落
var firstPara = doc.GetElement(0)

// 创建新段落
var newPara = Api.CreateParagraph()

// ============================================
// 3. 操作段落内容
// ============================================
// 添加文字
newPara.AddText('这是使用 JsAPI 添加的文字')

// 设置段落对齐方式
newPara.SetJc('center') // center, left, right, both, etc.

// 将段落添加到文档
doc.Push(newPara)

// ============================================
// 4. 获取当前选区
// ============================================
var range = doc.GetRange()

// 在选区位置添加批注
var comment = range.AddComment('这是批注内容')
comment.SetAuthor('JsAPI User')
comment.SetDisplayName('JsAPI User')

// ============================================
// 5. 获取所有批注
// ============================================
var comments = doc.GetComments()
for (var i = 0; i < comments.length; i++) {
  var c = comments[i]
  Api.GetConsole().Log('批注 ' + (i + 1) + ': ' + c.GetText())
}

// ============================================
// 6. 搜索文字
// ============================================
var searchText = '要搜索的文字'
var paragraphs = doc.GetElements(Api.c_oAscElementTypeParagraph)

for (var i = 0; i < paragraphs.length; i++) {
  var para = paragraphs[i]
  var paraText = para.GetText()
  
  if (paraText && paraText.indexOf(searchText) !== -1) {
    // 找到匹配的文字
    var paraRange = para.GetRange()
    var startPos = paraText.indexOf(searchText)
    
    // 创建范围定位找到的文字
    var foundRange = paraRange.Clone()
    foundRange.SetStart(paraRange.GetStart() + startPos)
    foundRange.SetEnd(paraRange.GetStart() + startPos + searchText.length)
    
    // 选中并滚动到该位置
    foundRange.Select()
    Api.ScrollToRange(foundRange)
    
    Api.GetConsole().Log('找到文字在第 ' + (i + 1) + ' 段')
    break
  }
}

// ============================================
// 7. 操作表格
// ============================================
// 创建表格
var table = Api.CreateTable(3, 3) // 3行3列

// 获取表格单元格
var cell = table.GetCell(0, 0)
var cellPara = cell.GetElement(0)
cellPara.AddText('单元格内容')

// 将表格添加到文档
doc.Push(table)

// ============================================
// 8. 操作图片
// ============================================
// 在段落中插入图片
var imagePara = Api.CreateParagraph()
var image = Api.CreateImage('https://example.com/image.png')
imagePara.AddDrawing(image)
doc.Push(imagePara)

// ============================================
// 9. 设置样式
// ============================================
var styledPara = Api.CreateParagraph()
styledPara.AddText('粗体文字')
var run = styledPara.GetElement(0)
run.SetBold(true)
run.SetFontSize(16)
doc.Push(styledPara)

// ============================================
// 10. 获取文档信息
// ============================================
var elementCount = doc.GetElementsCount()
Api.GetConsole().Log('文档共有 ' + elementCount + ' 个元素')

// ============================================
// 在插件中使用这些 JsAPI 代码
// ============================================
// 插件代码示例：
/*
window.Asc.plugin.button = function (id) {
  if (id === 'myButton') {
    this.callCommand(function () {
      // 在这里使用上面的 JsAPI 代码
      var doc = Api.GetDocument()
      var para = Api.CreateParagraph()
      para.AddText('通过插件按钮触发的 JsAPI 操作')
      doc.Push(para)
    }, true)
  }
}
*/

