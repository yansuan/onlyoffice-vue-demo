
(function (window, undefined) {
  window.Asc.plugin.init = function () {
    $('#paragraphs-btn').on('click', function () {

      window.Asc.plugin.callCommand(function () {
        var doc = Api.GetDocument();          // 获取当前文档
        var paragraphs = doc.GetAllParagraphs(); // 官方提供的段落获取方法

        for (var i = 0; i < paragraphs.length; i++) {
          var p = paragraphs[i];

          // 获取段落纯文本
          var text = p.GetText();

          p.Select()
          var currentPage = doc.GetCurrentPage();

          let searchResults = p.Search("text", false);
          if (searchResults.length > 0) {
            searchResults[0].Select();
            searchResults[0].SetBold(true);

            let deliveryRun = searchResults[0]
            var c = Api.AddComment(deliveryRun, "Align with new SLA policy? Current policy requires 5 days minimum.", "Operations Manager");

            // 获取段落的 range
            var range = searchResults[0].GetRange();

            if (range) {
              // console.log(range);
              // // 页码（从 0 开始）
              // var startPageIndex = range.GetStartPage();
              // var endPageIndex = range.GetEndPage();

              // // 段落在页面上的位置坐标
              // var startPos = range.GetStartPos(); // {Left, Top, Right, Bottom, Width, Height}
              // var endPos = range.GetEndPos();

              // console.log("段落文字:", text)
              // console.log("开始页码:", startPageIndex + 1);
              // console.log("结束页码:", endPageIndex + 1);
              // console.log("开始位置:", startPos);
              // console.log("结束位置:", endPos);

              let section = p.GetSection();
              let startPageNumber = section.GetStartPageNumber();
              console.log("段落文字:", text)
              console.log("开始页码:", startPageNumber);
              console.log("当前页码:", currentPage)
            }
          }
        }
      });
    });

    $("#pop1-btn").on('click', function () {
      window.Asc.plugin.showWindow(true, "pop1");
    })

  };

  window.Asc.plugin.button = function (id) {
    console.log("button:", id)
    // 只关闭按钮才关闭插件，其他按钮不关闭
    if (id === Asc.c_oAscPlugInButtonNames.close) {
      this.executeCommand("close", "");
    }
  };
})(window, undefined);
