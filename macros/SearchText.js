// 搜索文字宏 - 使用 JsAPI
// 使用方法：在 OnlyOffice 编辑器中，打开"开发工具" -> "宏" -> 新建宏，粘贴此代码并运行
// 修改 searchText 变量为要搜索的文字

(function () {
  var doc = Api.GetDocument()
  var searchText = '示例文字' // 修改这里为要搜索的文字
  var results = []

  if (!searchText || searchText.trim() === '') {
    Api.GetConsole().Log('搜索文字不能为空')
    return
  }

  // 遍历文档内容查找文字
  var paragraphs = doc.GetElements(Api.c_oAscElementTypeParagraph)

  for (var i = 0; i < paragraphs.length; i++) {
    var para = paragraphs[i]
    var paraText = para.GetText()

    if (paraText && paraText.indexOf(searchText) !== -1) {
      // 找到匹配的文字
      var range = para.GetRange()
      var startPos = paraText.indexOf(searchText)

      // 创建范围来定位找到的文字
      var foundRange = range.Clone()
      foundRange.SetStart(range.GetStart() + startPos)
      foundRange.SetEnd(range.GetStart() + startPos + searchText.length)

      results.push({
        paragraphIndex: i,
        startPos: startPos,
        text: paraText.substring(Math.max(0, startPos - 20), startPos + searchText.length + 20),
      })
    }
  }

  // 显示搜索结果
  if (results.length > 0) {
    var message =
      '找到 ' +
      results.length +
      ' 处匹配\n' +
      '位置: ' +
      results
        .map(function (r) {
          return '第' + (r.paragraphIndex + 1) + '段'
        })
        .join(', ')

    Api.GetConsole().Log(message)

    // 导航到第一个结果并高亮显示
    var firstResult = results[0]
    var firstPara = paragraphs[firstResult.paragraphIndex]
    var firstRange = firstPara.GetRange()
    var foundRange = firstRange.Clone()
    foundRange.SetStart(firstRange.GetStart() + firstResult.startPos)
    foundRange.SetEnd(firstRange.GetStart() + firstResult.startPos + searchText.length)

    // 选中找到的文字
    foundRange.Select()

    // 滚动到该位置
    Api.ScrollToRange(foundRange)

    Api.GetConsole().Log('已定位到第 ' + (firstResult.paragraphIndex + 1) + ' 段')
  } else {
    Api.GetConsole().Log('未找到匹配的文字: ' + searchText)
  }
})()

