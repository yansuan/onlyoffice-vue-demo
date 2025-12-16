// 添加批注宏 - 使用 JsAPI
// 使用方法：在 OnlyOffice 编辑器中，打开"开发工具" -> "宏" -> 新建宏，粘贴此代码并运行

(function () {
  var doc = Api.GetDocument()
  var range = doc.GetRange() // 当前光标位置或选中的范围

  // 要添加的批注内容
  var commentText = '这是通过 JsAPI 添加的批注\n时间: ' + new Date().toLocaleString()
  var author = 'MacroUser'

  // 在当前位置添加批注
  var comment = range.AddComment(commentText)

  // 设置批注作者
  comment.SetAuthor(author)
  comment.SetDisplayName(author)

  // 显示成功消息
  Api.GetConsole().Log('批注已添加: ' + commentText)
})()

