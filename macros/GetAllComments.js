// 获取所有批注宏 - 使用 JsAPI
// 使用方法：在 OnlyOffice 编辑器中，打开"开发工具" -> "宏" -> 新建宏，粘贴此代码并运行

(function () {
  var doc = Api.GetDocument()
  var comments = doc.GetComments() // 获取所有批注
  var result = []

  // 遍历所有批注
  for (var i = 0; i < comments.length; i++) {
    var c = comments[i]
    result.push({
      index: i + 1,
      author: c.GetAuthor(), // 批注作者
      text: c.GetText(), // 批注内容
      date: c.GetDateTime(), // 批注时间
    })
  }

  // 输出到控制台
  var json = JSON.stringify(result, null, 2)
  Api.GetConsole().Log('批注列表:\n' + json)

  // 也可以将结果追加到文档末尾
  var para = doc.GetEnd().GetOrCreatePara()
  para.AddText('批注列表（共 ' + result.length + ' 条）:\n')
  para.AddText(json)
  para.AddText('\n\n')

  // 显示消息
  if (result.length > 0) {
    Api.GetConsole().Log('找到 ' + result.length + ' 条批注')
  } else {
    Api.GetConsole().Log('文档中没有批注')
  }
})()

