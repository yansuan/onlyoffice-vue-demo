window.Asc.plugin.init = function () {
  // no UI to init
}

window.Asc.plugin.button = function (id) {
  if (id === Asc.c_oAscPlugInButtonNames.close) {
    this.executeMethod('Close', [])
    return
  }

  if (id === 'addComment') {
    addComment.call(this)
  } else if (id === 'getComments') {
    dumpComments.call(this)
  }
}

function addComment() {
  this.callCommand(function () {
    var doc = Api.GetDocument()
    var range = doc.GetRange()
    var comment = range.AddComment('JsAPI 插件添加的批注')
    comment.SetAuthor('PluginUser')
    comment.SetDisplayName('PluginUser')
  }, true)
}

function dumpComments() {
  this.callCommand(function () {
    var doc = Api.GetDocument()
    var comments = doc.GetComments()
    var info = []

    for (var i = 0; i < comments.length; i++) {
      var c = comments[i]
      info.push({
        index: i + 1,
        author: c.GetAuthor(),
        text: c.GetText(),
        date: c.GetDateTime(),
      })
    }

    var json = JSON.stringify(info, null, 2)
    Api.GetConsole().Log('批注列表:\\n' + json)
    var para = doc.GetEnd().GetOrCreatePara()
    para.AddText('批注列表：\\n' + json)
  }, true)
}

