window.Asc.plugin.init = function () {
  // 存储搜索状态
  this.searchText = ''
  this.searchResults = []
  this.currentIndex = -1
  this.lastRequestId = null

  // 启动监听器，检查外部搜索请求
  var plugin = this
  this.checkInterval = setInterval(function () {
    checkExternalSearchRequest.call(plugin)
  }, 200) // 每 200ms 检查一次
}

window.Asc.plugin.onExternalMouseUp = function () {
  // 插件关闭时清理
  if (this.checkInterval) {
    clearInterval(this.checkInterval)
  }
}

// 检查外部搜索请求
function checkExternalSearchRequest() {
  var plugin = this
  try {
    var requestStr = localStorage.getItem('onlyoffice_search_request')
    if (requestStr) {
      var request = JSON.parse(requestStr)
      // 检查是否是新的请求
      if (request.id !== plugin.lastRequestId) {
        plugin.lastRequestId = request.id
        plugin.searchText = request.text
        plugin.currentIndex = -1
        // 清除请求，避免重复处理
        localStorage.removeItem('onlyoffice_search_request')
        // 执行搜索
        performSearch.call(plugin, request.id)
      }
    }
  } catch (e) {
    // 忽略解析错误
  }
}

window.Asc.plugin.button = function (id) {
  if (id === Asc.c_oAscPlugInButtonNames.close) {
    this.executeMethod('Close', [])
    return
  }

  if (id === 'searchText') {
    searchText.call(this)
  } else if (id === 'findNext') {
    findNext.call(this)
  }
}

function searchText() {
  var plugin = this

  // 显示输入对话框
  plugin.showDialog(
    'https://onlyoffice.github.io/sdkjs-plugins/v1/resources/dialogs/input.html',
    function (data) {
      if (data && data.text) {
        plugin.searchText = data.text
        plugin.currentIndex = -1
        performSearch.call(plugin, null) // 内部触发，不需要 requestId
      }
    },
    true,
    {
      width: 400,
      height: 150,
    },
    {
      text: '请输入要搜索的文字:',
      value: plugin.searchText || '',
    },
  )
}

function findNext() {
  var plugin = this
  if (plugin.searchResults.length === 0) {
    plugin.showInfo('请先执行搜索')
    return
  }
  plugin.currentIndex = (plugin.currentIndex + 1) % plugin.searchResults.length
  navigateToResult.call(plugin, plugin.currentIndex)
}

function performSearch(requestId) {
  var plugin = this
  plugin.callCommand(
    function () {
      var doc = Api.GetDocument()
      var searchText = plugin.searchText
      var results = []

      if (!searchText || searchText.trim() === '') {
        Api.GetConsole().Log('搜索文字不能为空')
        // 返回空结果
        if (requestId) {
          localStorage.setItem(
            'onlyoffice_search_result',
            JSON.stringify({
              requestId: requestId,
              found: false,
              count: 0,
              locations: [],
            }),
          )
        }
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
            range: foundRange,
            text: paraText,
          })
        }
      }

      plugin.searchResults = results
      plugin.currentIndex = -1

      // 将结果写回 localStorage（供外部 Vue 组件读取）
      if (requestId) {
        var locations = results.map(function (r) {
          return '第' + (r.paragraphIndex + 1) + '段'
        })

        localStorage.setItem(
          'onlyoffice_search_result',
          JSON.stringify({
            requestId: requestId,
            found: results.length > 0,
            count: results.length,
            locations: locations,
          }),
        )
      }

      if (results.length > 0) {
        var message =
          '找到 ' +
          results.length +
          ' 处匹配\n' +
          '段落位置: ' +
          results.map(function (r) {
            return '第' + (r.paragraphIndex + 1) + '段'
          }).join(', ')

        Api.GetConsole().Log(message)

        // 导航到第一个结果
        if (results.length > 0) {
          navigateToResult.call(plugin, 0)
        }
      } else {
        Api.GetConsole().Log('未找到匹配的文字: ' + searchText)
        plugin.showInfo('未找到匹配的文字: ' + searchText)
      }
    },
    true,
  )
}

function navigateToResult(index) {
  var plugin = this
  if (index < 0 || index >= plugin.searchResults.length) {
    return
  }

  plugin.callCommand(
    function () {
      var result = plugin.searchResults[index]
      var range = result.range

      // 选中找到的文字
      range.Select()

      // 滚动到该位置
      Api.ScrollToRange(range)

      // 高亮显示（通过选中实现）
      var message =
        '找到第 ' +
        (index + 1) +
        ' / ' +
        plugin.searchResults.length +
        ' 处\n' +
        '位置: 第' +
        (result.paragraphIndex + 1) +
        '段\n' +
        '文字: ' +
        result.text.substring(Math.max(0, result.startPos - 20), result.startPos + plugin.searchText.length + 20)

      Api.GetConsole().Log(message)
    },
    true,
  )
}

