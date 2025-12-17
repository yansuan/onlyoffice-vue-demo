// WebSocket 客户端 - 用于接收外部命令并执行 JsAPI
(function (window) {
  'use strict'

  let ws = null
  let reconnectTimer = null
  const handledRequests = new Set()
  const WS_URL = 'ws://192.168.93.1:4000?type=plugin'
  const RECONNECT_DELAY = 3000

  // 连接 WebSocket
  function connect() {
    try {
      ws = new WebSocket(WS_URL)

      ws.onopen = function () {
        console.log('[WebSocket] 插件已连接到服务器')
        if (reconnectTimer) {
          clearTimeout(reconnectTimer)
          reconnectTimer = null
        }
      }

      ws.onmessage = function (event) {
        try {
          const data = JSON.parse(event.data)
          console.log('[WebSocket] 收到命令:', data)

          // 处理命令
          handleCommand(data)
        } catch (error) {
          console.error('[WebSocket] 消息解析错误:', error)
        }
      }

      ws.onerror = function (error) {
        console.error('[WebSocket] 连接错误:', error)
      }

      ws.onclose = function () {
        console.log('[WebSocket] 连接已关闭，尝试重连...')
        // 自动重连
        reconnectTimer = setTimeout(connect, RECONNECT_DELAY)
      }
    } catch (error) {
      console.error('[WebSocket] 连接失败:', error)
      reconnectTimer = setTimeout(connect, RECONNECT_DELAY)
    }
  }

  // 发送结果给 Vue 应用
  function sendResult(requestId, type, result, error) {
    if (ws && ws.readyState === WebSocket.OPEN) {
      const response = {
        requestId: requestId,
        type: type,
        result: result,
        error: error,
      }
      ws.send(JSON.stringify(response))
      console.log(
        '[WebSocket] 发送响应 RequestId:',
        requestId,
        '类型:',
        type,
        'error:',
        error,
        'resultSummary:',
        result && typeof result === 'object' ? Object.keys(result) : typeof result
      )
    } else {
      console.error('[WebSocket] 连接未就绪，无法发送结果 RequestId:', requestId)
    }
  }

  // 处理命令
  function handleCommand(data) {
    const requestId = data.requestId
    const { command, params } = data

    if (!requestId) {
      console.error('[WebSocket] 收到无 RequestId 的命令:', data)
      return
    }

    console.log('[WebSocket] 收到命令 RequestId:', requestId, '命令:', command, 'params:', params)

    if (handledRequests.has(requestId)) {
      console.warn('[WebSocket] 检测到重复命令, 忽略执行. RequestId:', requestId, 'command:', command)
      return
    }
    handledRequests.add(requestId)

    if (!window.Asc || !window.Asc.plugin) {
      sendResult(requestId, 'error', null, '插件环境未初始化')
      return
    }

    // 在插件端不做超时控制，由发送端处理超时逻辑
    // 这里只负责执行命令，并返回成功或错误结果
    try {
      switch (command) {
        case 'insertText':
          // 插入文本到文档末尾
          Asc.scope.text = params.text || '未设置'
          Asc.scope.requestId = requestId
          Asc.scope.sendResult = sendResult
          console.log(
            '[Plugin] insertText 开始执行, RequestId:',
            requestId,
            'text:',
            Asc.scope.text
          )
          console.log('[Plugin] callCommand:', window.Asc.plugin.callCommand);
          window.Asc.plugin.callCommand(
            () => {
              try {
                const oDocument = Api.GetDocument()
                const oParagraph = Api.CreateParagraph()
                oParagraph.AddText(Asc.scope.text)
                oDocument.InsertContent([oParagraph])
                return 'success'
              } catch (error) {
                return error.message
              }
            },
            false,
            true,
            (returnValue) => {
              if (returnValue === 'success') {
                sendResult(requestId, 'success', { message: returnValue }, null)
              } else {
                sendResult(requestId, 'error', { message: returnValue }, null)
              }
            }
          )
          break

        case 'searchText':
          // 搜索文本并读取页码和行数
          Asc.scope.text = params.text || ''
          window.Asc.plugin.callCommand(
            () => {
              try {
                const oDocument = Api.GetDocument()
                const oSearchResults = oDocument.Search(Asc.scope.text, false)
                // oSearchResults = oDocument.Search(Asc.scope.text, false)
                const results = []
                for (let i = 0; i < oSearchResults.length; i++) {
                  const oSearchResult = oSearchResults[i]
                  // oSearchResult.SetBold(true)
                  // Try to get line numbr using available methods
                  // let lineNumber = null;
                  // try {
                  //   // Try different methods that might give us position information
                  //   if (oSearchResult.GetStartPos) {
                  //     const pos = oSearchResult.GetStartPos();
                  //     console.log('GetStartPos():', pos);
                  //     // If it's a number, it might be the line number
                  //     if (typeof pos === 'number') {
                  //       lineNumber = pos + 1; // +1 to make it 1-based
                  //     }
                  //   }
                  // } catch (e) {
                  //   console.error('Error getting line number:', e);
                  // }
                  results.push({
                    text: oSearchResult.GetText(),
                    page: oSearchResult.GetStartPage() + 1,
                    // line: lineNumber,
                    // pageNumber: {
                    //   start: oSearchResult ? oSearchResult.GetStartPage() : null,
                    //   end: oSearchResult ? oSearchResult.GetEndPage() : null,
                    // },
                    // pos: {
                    //   start: oSearchResult ? oSearchResult.GetStartPos() : null,
                    //   end: oSearchResult ? oSearchResult.GetEndPos() : null,
                    // }
                  })
                }

                console.log('[Plugin] searchText results', results);

                return results
              } catch (error) {
                return error.message
              }
            },
            false,
            false,
            (returnValue) => {
              sendResult(requestId, 'success', returnValue, null)

            }
          )
          break

        case 'getAllComments':
          // 获取所有批注
          window.Asc.plugin.executeMethod(
            'GetAllComments',
            [],
            function (comments) {
              sendResult(requestId, 'success', { comments: comments }, null)
            }
          )
          break

        case 'addComment':

          // 搜索文本并添加批注
          Asc.scope.text = params.text || ''
          console.log('[Plugin] addComment 开始执行, RequestId:', requestId, 'text:', Asc.scope.text)
          window.Asc.plugin.callCommand(
            () => {
              try {
                const oDocument = Api.GetDocument()
                const oSearchResults = oDocument.Search(Asc.scope.text, false)
                const results = []
                for (let i = 0; i < oSearchResults.length; i++) {
                  const oSearchResult = oSearchResults[i]
                  oSearchResult.SetBold(true)

                  const range = oSearchResult.GetRange()
                  if (!range) {
                    continue
                  }

                  const oComment = range.AddComment('这是批注信息.', '操作用户')

                  results.push({
                    comment: oComment,
                    text: oSearchResult.GetText(),
                    page: range ? range.GetStartPage() + 1 : null,
                    position: {
                      left: range ? range.GetStartPos().Left : null,
                      top: range ? range.GetStartPos().Top : null,
                    },
                  })
                }

                console.log('[Plugin] addComment results', results)

                return 'success'
              } catch (error) {
                return error.message
              }
            },
            false,
            false,
            (returnValue) => {
              console.log('[Plugin] addComment 执行完成, RequestId:', requestId, 'returnValue:', returnValue)
              if (returnValue === 'success') {
                sendResult(requestId, 'success', { message: returnValue }, null)
              } else {
                sendResult(requestId, 'error', { message: returnValue }, null)
              }
            }
          )
          break

        case 'addCommentToSelection':
          // 在编辑器当前选中的内容上添加批注
          console.log('[Plugin] addCommentToSelection 开始执行, RequestId:', requestId)
          const commentText = params.text || params.commentText || ''
          const userName = params.userName || params.user || '审阅者'

          if (!commentText.trim()) {
            sendResult(requestId, 'error', null, '批注内容不能为空')
            return
          }

          Asc.scope.commentText = commentText
          Asc.scope.userName = userName
          Asc.scope.requestId = requestId
          Asc.scope.sendResult = sendResult

          window.Asc.plugin.callCommand(
            () => {
              try {
                const doc = Api.GetDocument()
                // 获取当前选中的内容
                const oRange = doc.GetRangeBySelect()
                // const oRange = doc.GetCurrentRun()

                if (!oRange) {
                  return {
                    success: false,
                    error: '未选中任何内容，请先选择要添加批注的文本'
                  }
                }

                const pageNumber = oRange.GetStartPage() + 1;

                // 获取选中的文本内容
                const selectedText = oRange.GetText()
                if (!selectedText || selectedText.trim().length === 0) {
                  return {
                    success: false,
                    error: '选中的内容为空，请选择包含文本的内容'
                  }
                }

                const commentText = Asc.scope.commentText
                const userName = Asc.scope.userName


                // 在选中的内容上添加批注
                const oComment = oRange.AddComment(commentText, userName)

                // 获取批注和选中内容的位置信息
                const commentInfo = {
                  commentId: oComment ? oComment.GetId() : null,
                  selectedText: selectedText,
                  commentText: commentText,
                  userName: userName,
                  page: pageNumber,
                }

                console.log('[Plugin] addCommentToSelection 成功, RequestId:', Asc.scope.requestId, 'commentInfo:', commentInfo)

                return {
                  success: true,
                  ...commentInfo
                }
              } catch (error) {
                console.error('[Plugin] addCommentToSelection 执行异常:', error)
                return {
                  success: false,
                  error: error.message || '添加批注时发生错误'
                }
              }
            },
            false,
            false, // 重新计算文档以显示批注
            (returnValue) => {
              const currentRequestId = Asc.scope.requestId
              const currentSendResult = Asc.scope.sendResult

              console.log('[Plugin] addCommentToSelection 执行完成, RequestId:', currentRequestId, 'returnValue:', returnValue)

              if (returnValue && returnValue.success) {
                currentSendResult(currentRequestId, 'success', returnValue, null)
              } else {
                const errorMsg = returnValue && returnValue.error
                  ? returnValue.error
                  : '添加批注失败'
                currentSendResult(currentRequestId, 'error', null, errorMsg)
              }
            }
          )
          break

        case 'getAllParagraphs':
          // 获取所有段落
          console.log('[Plugin] getAllParagraphs 开始执行, RequestId:', requestId)
          window.Asc.plugin.callCommand(
            function () {
              var doc = Api.GetDocument()
              var paragraphs = doc.GetAllParagraphs()
              var results = []
              var lastPageIndex = -1;
              var currentRow = 0;
              for (var i = 0; i < paragraphs.length; i++) {
                var p = paragraphs[i]
                var text = p.GetText()
                var range = p.GetRange()
                var currentPage = range ? range.GetStartPage() + 1 : -2;
                if (currentPage !== lastPageIndex) {
                  lastPageIndex = currentPage;
                  currentRow = 0;
                } else {
                  currentRow++;
                }
                results.push({
                  index: i,
                  text: text,
                  page: currentPage,
                  line: currentRow,
                  pageNumber: {
                    start: range ? range.GetStartPage() : null,
                    end: range ? range.GetEndPage() : null,
                  },
                  pos: {
                    start: range ? range.GetStartPos() : null,
                    end: range ? range.GetEndPos() : null,
                  }
                })
              }
              console.log('[Plugin] getAllParagraphs 结果:', results);
              return results
            },
            false,
            false,
            function (result) {
              console.log(
                '[Plugin] getAllParagraphs 执行完成, RequestId:',
                requestId,
                '段落数量:',
                result ? result.length : 0
              )
              sendResult(requestId, 'success', { paragraphs: result }, null)
            }
          )
          break

        case 'saveDocument':
          // 保存文档：通过 Command Service API 的 forcesave 命令
          // forcesave 会自动保存文档并触发 callback，不需要先调用 Api.Save()
          console.log('[Plugin] saveDocument 开始执行, RequestId:', requestId)
          console.log('[Plugin] 收到参数:', params)

          try {
            const documentKey = params.documentKey
            if (!documentKey) {
              throw new Error('缺少 documentKey 参数')
            }

            // 调用 Command Service API 的 forcesave 命令
            // 这个命令会：1) 保存文档到 OnlyOffice 服务器  2) 触发 callback (status=6)
            const commandServiceUrl = 'http://192.168.93.128:8101/coauthoring/CommandService.ashx'

            const commandPayload = {
              c: 'forcesave',
              key: documentKey,
              userdata: 'manual-save-from-button'
            }

            console.log('[Plugin] 发送 forcesave 命令, key:', documentKey)
            console.log('[Plugin] Command Service URL:', commandServiceUrl)

            fetch(commandServiceUrl, {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
              },
              body: JSON.stringify(commandPayload)
            })
              .then(response => response.json())
              .then(result => {
                console.log('[Plugin] Command Service 响应:', result)

                if (result.error === 0) {
                  console.log('[Plugin] ✓ forcesave 命令发送成功')
                  console.log('[Plugin] ✓ OnlyOffice 将触发 callback (status=6) 并保存文档')
                  sendResult(requestId, 'success', {
                    success: true,
                    message: '文档保存已触发',
                    method: 'Command Service API forcesave',
                    note: 'OnlyOffice 将通过 callbackUrl 发送 status=6 回调并保存文档到服务器',
                    commandServiceResponse: result
                  }, null)
                } else {
                  console.error('[Plugin] ✗ forcesave 命令失败:', result)
                  sendResult(requestId, 'error', null, 'Command Service 返回错误: ' + result.error)
                }
              })
              .catch(error => {
                console.error('[Plugin] Command Service 请求失败:', error)
                sendResult(requestId, 'error', null, 'forcesave 请求失败: ' + error.message)
              })
          } catch (error) {
            console.error('[Plugin] saveDocument 执行异常, RequestId:', requestId, 'error:', error)
            sendResult(requestId, 'error', null, '保存失败: ' + error.message)
          }
          break

        case 'downloadDocument':
          // 下载正在编辑的文档并保存到服务器
          console.log('[Plugin] downloadDocument 开始执行, RequestId:', requestId)
          console.log('[Plugin] 收到参数:', params)

          try {
            const documentKey = params.documentKey
            if (!documentKey) {
              throw new Error('缺少 documentKey 参数')
            }

            console.log('[Plugin] 步骤1: 调用 Api.Save() 保存当前编辑内容')

            // 先保存当前编辑内容到内存
            window.Asc.plugin.callCommand(
              function () {
                Api.Save()
                return { success: true }
              },
              false,
              false,
              function () {
                console.log('[Plugin] Api.Save() 完成')
                console.log('[Plugin] 步骤2: 调用 forcesave 触发文档保存')

                // 等待 1 秒后调用 forcesave
                setTimeout(function () {
                  const commandServiceUrl = 'http://192.168.93.128:8101/coauthoring/CommandService.ashx'

                  // 调用 forcesave 触发保存
                  fetch(commandServiceUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                      c: 'forcesave',
                      key: documentKey,
                      userdata: 'download-document'  // 标记为下载操作
                    })
                  })
                    .then(response => response.json())
                    .then(forcesaveResult => {
                      console.log('[Plugin] forcesave 响应:', forcesaveResult)

                      if (forcesaveResult.error === 0) {
                        console.log('[Plugin] ✓ forcesave 成功')
                        console.log('[Plugin] 步骤3: 等待 callback 处理文档下载')
                        console.log('[Plugin] OnlyOffice 将通过 callback 发送文档到服务器')

                        // 等待 2 秒后再获取文档信息
                        setTimeout(function () {
                          console.log('[Plugin] 步骤4: 调用 info 获取文档地址')

                          fetch(commandServiceUrl, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ c: 'info', key: documentKey })
                          })
                            .then(response => response.json())
                            .then(infoResult => {
                              console.log('[Plugin] info 响应:', infoResult)

                              if (infoResult.error === 0 && infoResult.url) {
                                console.log('[Plugin] ✓ 获取到文档地址:', infoResult.url)
                                console.log('[Plugin] 步骤5: 下载文档并发送到服务器')

                                // 下载文档
                                fetch(infoResult.url)
                                  .then(response => response.arrayBuffer())
                                  .then(arrayBuffer => {
                                    // 转换为 base64
                                    const bytes = new Uint8Array(arrayBuffer)
                                    let binary = ''
                                    for (let i = 0; i < bytes.byteLength; i++) {
                                      binary += String.fromCharCode(bytes[i])
                                    }
                                    const base64 = btoa(binary)

                                    console.log('[Plugin] ✓ 文档下载完成, 大小:', (arrayBuffer.byteLength / 1024).toFixed(2), 'KB')

                                    // 发送到后端保存
                                    return fetch('http://192.168.93.1:4000/api/download-document', {
                                      method: 'POST',
                                      headers: { 'Content-Type': 'application/json' },
                                      body: JSON.stringify({
                                        fileData: base64,
                                        fileName: 'downloaded_document.docx',
                                        timestamp: new Date().toISOString()
                                      })
                                    })
                                  })
                                  .then(response => response.json())
                                  .then(saveResult => {
                                    console.log('[Plugin] ✓ 文档已保存到服务器:', saveResult)
                                    sendResult(requestId, 'success', {
                                      success: true,
                                      message: '文档已成功下载到服务器',
                                      filePath: saveResult.filePath,
                                      fileName: saveResult.fileName,
                                      fileSize: saveResult.fileSize
                                    }, null)
                                  })
                                  .catch(error => {
                                    console.error('[Plugin] 下载或保存失败:', error)
                                    sendResult(requestId, 'error', null, '下载失败: ' + error.message)
                                  })
                              } else {
                                console.error('[Plugin] ✗ info 返回错误:', infoResult.error)
                                sendResult(requestId, 'error', null, 'info 命令返回错误: ' + infoResult.error)
                              }
                            })
                            .catch(error => {
                              console.error('[Plugin] info 请求失败:', error)
                              sendResult(requestId, 'error', null, 'info 请求失败: ' + error.message)
                            })
                        }, 2000)
                      } else {
                        console.error('[Plugin] ✗ forcesave 失败, error:', forcesaveResult.error)
                        sendResult(requestId, 'error', null, 'forcesave 返回错误: ' + forcesaveResult.error)
                      }
                    })
                    .catch(error => {
                      console.error('[Plugin] forcesave 请求失败:', error)
                      sendResult(requestId, 'error', null, 'forcesave 请求失败: ' + error.message)
                    })
                }, 1000)
              }
            )
          } catch (error) {
            console.error('[Plugin] downloadDocument 执行异常, RequestId:', requestId, 'error:', error)
            sendResult(requestId, 'error', null, '下载失败: ' + error.message)
          }
          break

        case 'executeJsAPI':
          // 执行自定义 JsAPI 代码（通过函数字符串）
          // 注意：这是一个简化的实现，实际使用中需要更安全的代码执行方式
          try {
            var func = new Function('Api', params.code || '')
            console.log(
              '[Plugin] executeJsAPI 开始执行, RequestId:',
              requestId,
              'code 长度:',
              (params.code || '').length
            )
            var shouldRecalc =
              typeof params.recalc === 'boolean' ? params.recalc : true
            window.Asc.plugin.callCommand(
              function () {
                return func(Api)
              },
              false,
              shouldRecalc,
              function (result) {
                console.log(
                  '[Plugin] executeJsAPI 执行完成, RequestId:',
                  requestId,
                  'resultType:',
                  typeof result
                )
                sendResult(requestId, 'success', { result: result }, null)
              }
            )
          } catch (error) {
            console.error(
              '[Plugin] executeJsAPI 执行异常, RequestId:',
              requestId,
              'error:',
              error
            )
            sendResult(requestId, 'error', null, error.message)
            break
          }
          break

        case 'scrollToPage':
          // 滚动到指定页码
          console.log('[Plugin] scrollToPage 开始执行, RequestId:', requestId, 'pageNumber:', params.pageNumber);
          const pageNumber = parseInt(params.pageNumber) || 1;
          Asc.scope.pageNumber = pageNumber;

          window.Asc.plugin.callCommand(
            function () {
              try {
                // 使用正确的 API 方法跳转到指定页面
                console.log('[Plugin] 正在跳转到第', Asc.scope.pageNumber, '页');
                const doc = Api.GetDocument()
                const isOk = doc.GoToPage(Asc.scope.pageNumber); // API 中页码从 0 开始
                console.log('[Plugin] 使用 asc_GoToPage 跳转完成,ok:', isOk);

              } catch (error) {
                console.error('[Plugin] scrollToPage 执行错误:', error);
                return {
                  success: false,
                  error: error.message || '滚动到指定页失败',
                  details: error
                };
              }
            },
            false,
            true,
            function (result) {
              window.Asc.plugin.callCommand(
                function () {
                  try {
                    // 使用正确的 API 方法跳转到指定页面
                    console.log('[Plugin] 正在跳转到第', Asc.scope.pageNumber, '页');
                    const doc = Api.GetDocument()
                    const isOk = doc.GoToPage(Asc.scope.pageNumber - 1); // API 中页码从 0 开始
                    console.log('[Plugin] 使用 asc_GoToPage 跳转完成,ok:', isOk);
                    return {
                      success: true,
                      currentPage: Asc.scope.pageNumber
                    };

                  } catch (error) {
                    console.error('[Plugin] scrollToPage 执行错误:', error);
                    return {
                      success: false,
                      error: error.message || '滚动到指定页失败',
                      details: error
                    };
                  }
                },
                false,
                true,
                function (result) {
                  console.log('[Plugin] scrollToPage 执行完成, RequestId:', requestId, 'result:', result);
                  if (result && result.success) {
                    sendResult(requestId, 'success', {
                      message: `已滚动到第 ${result.currentPage} 页`,
                      currentPage: result.currentPage
                    });
                  } else {
                    console.error('[Plugin] scrollToPage 执行失败:', result?.error);
                    sendResult(requestId, 'error', {
                      error: '滚动到指定页失败',
                      details: result?.error || '未知错误'
                    });
                  }
                }
              );
            }
          );
          break

        case 'searchAndNavigate':
          // 搜索文本并跳转到对应页面，高亮显示
          console.log('[Plugin] searchAndNavigate 开始执行, RequestId:', requestId);
          const searchText = params.searchText || '';
          const targetPage = params.pageNumber;

          if (!searchText) {
            sendResult(requestId, 'error', null, '搜索内容不能为空');
            break;
          }

          if (!targetPage || targetPage < 1) {
            sendResult(requestId, 'error', null, '页码是必传值，且必须大于0');
            break;
          }

          Asc.scope.searchText = searchText;
          Asc.scope.targetPage = targetPage;
          Asc.scope.requestId = requestId;
          Asc.scope.sendResult = sendResult;

          window.Asc.plugin.callCommand(
            function () {
              try {
                const oDocument = Api.GetDocument();

                // 使用 Search 方法搜索文本
                const oSearchResults = oDocument.Search(Asc.scope.searchText, false);
                const pageResults = [];

                // 只处理指定页面的搜索结果
                for (let i = 0; i < oSearchResults.length; i++) {
                  const oSearchResult = oSearchResults[i];
                  const page = oSearchResult.GetStartPage() + 1;

                  // 只保留当前页面的结果
                  if (page === Asc.scope.targetPage) {
                    pageResults.push({
                      text: oSearchResult.GetText(),
                      page: page,
                      index: i,
                      searchResult: oSearchResult
                    });
                  }
                }

                console.log('[Plugin] 在第', Asc.scope.targetPage, '页找到', pageResults.length, '个结果');

                if (pageResults.length === 0) {
                  // 在指定页面没有找到匹配内容
                  return {
                    success: false,
                    error: `在第 ${Asc.scope.targetPage} 页未找到匹配的内容"${Asc.scope.searchText}"`
                  };
                }

                // 先跳转到指定页面
                oDocument.GoToPage(Asc.scope.targetPage - 1);

                // 高亮所有匹配项
                for (let j = 0; j < pageResults.length; j++) {
                  const result = pageResults[j];
                  const oSearchResult = oSearchResults[result.index];
                  oSearchResult.SetHighlight("yellow");
                }

                console.log('[Plugin] 已跳转到第', Asc.scope.targetPage, '页，并高亮', pageResults.length, '个匹配项');

                return {
                  success: true,
                  results: pageResults.map(r => ({ page: r.page, text: r.text, index: r.index })),
                  count: pageResults.length,
                  targetPage: Asc.scope.targetPage
                };
              } catch (error) {
                console.error('[Plugin] searchAndNavigate 执行错误:', error);
                return {
                  success: false,
                  error: error.message || '搜索失败'
                };
              }
            },
            false,
            true,
            function (result) {
              console.log('[Plugin] searchAndNavigate 执行完成, RequestId:', Asc.scope.requestId, 'result:', result);
              if (result && result.success) {
                Asc.scope.sendResult(Asc.scope.requestId, 'success', {
                  results: result.results,
                  count: result.count,
                  targetPage: result.targetPage,
                  message: `在第 ${result.targetPage} 页找到 ${result.count} 个匹配项`
                }, null);
              } else {
                Asc.scope.sendResult(Asc.scope.requestId, 'error', null, result?.error || '搜索失败');
              }
            }
          );
          break

        case 'navigateToSearchResult':
          // 导航到指定页面的搜索结果（第一个匹配项）
          console.log('[Plugin] navigateToSearchResult 开始执行, RequestId:', requestId);
          const navSearchText = params.searchText || '';
          const navPageNumber = params.pageNumber;

          if (!navSearchText) {
            sendResult(requestId, 'error', null, '搜索内容不能为空');
            break;
          }

          if (!navPageNumber || navPageNumber < 1) {
            sendResult(requestId, 'error', null, '页码是必传值，且必须大于0');
            break;
          }

          Asc.scope.navSearchText = navSearchText;
          Asc.scope.navPageNumber = navPageNumber;
          Asc.scope.requestId = requestId;
          Asc.scope.sendResult = sendResult;

          window.Asc.plugin.callCommand(
            function () {
              try {
                const oDocument = Api.GetDocument();

                // 使用 Search 方法搜索文本
                const oSearchResults = oDocument.Search(Asc.scope.navSearchText, false);
                const pageResults = [];

                // 只处理指定页面的搜索结果
                for (let i = 0; i < oSearchResults.length; i++) {
                  const oSearchResult = oSearchResults[i];
                  const page = oSearchResult.GetStartPage() + 1;

                  // 只保留当前页面的结果
                  if (page === Asc.scope.navPageNumber) {
                    pageResults.push({
                      page: page,
                      index: i,
                      searchResult: oSearchResult
                    });
                  }
                }

                if (pageResults.length === 0) {
                  return {
                    success: false,
                    error: `在第 ${Asc.scope.navPageNumber} 页未找到匹配的内容"${Asc.scope.navSearchText}"`
                  };
                }

                // 先跳转到指定页面
                oDocument.GoToPage(Asc.scope.navPageNumber - 1);

                // 高亮所有匹配项
                for (let j = 0; j < pageResults.length; j++) {
                  const result = pageResults[j];
                  const oSearchResult = oSearchResults[result.index];
                  oSearchResult.SetHighlight("yellow");
                }

                console.log('[Plugin] 已跳转到第', Asc.scope.navPageNumber, '页，并高亮', pageResults.length, '个匹配项');

                // 获取该页面上的第一个结果
                const targetResult = pageResults[0];
                const oSearchResult = oSearchResults[targetResult.index];

                return {
                  success: true,
                  page: Asc.scope.navPageNumber,
                  pageResultCount: pageResults.length,
                  text: oSearchResult.GetText()
                };
              } catch (error) {
                console.error('[Plugin] navigateToSearchResult 执行错误:', error);
                return {
                  success: false,
                  error: error.message || '导航失败'
                };
              }
            },
            false,
            true,
            function (result) {
              console.log('[Plugin] navigateToSearchResult 执行完成, RequestId:', Asc.scope.requestId, 'result:', result);
              if (result && result.success) {
                Asc.scope.sendResult(Asc.scope.requestId, 'success', {
                  page: result.page,
                  pageResultCount: result.pageResultCount,
                  text: result.text,
                  message: `已跳转到第 ${result.page} 页的第一个匹配项（该页共 ${result.pageResultCount} 个）`
                }, null);
              } else {
                Asc.scope.sendResult(Asc.scope.requestId, 'error', null, result?.error || '导航失败');
              }
            }
          );
          break

        case 'clearHighlights':
          // 清除所有高亮
          console.log('[Plugin] clearHighlights 开始执行, RequestId:', requestId);

          // 获取之前搜索的文本和页码
          Asc.scope.clearRequestId = requestId;
          Asc.scope.clearSendResult = sendResult;
          Asc.scope.lastSearchText = params.searchText || '';
          Asc.scope.lastPageNumber = params.pageNumber || 0;

          window.Asc.plugin.callCommand(
            function () {
              try {
                const oDocument = Api.GetDocument();

                // 如果有搜索文本和页码，先搜索该页面的内容
                if (Asc.scope.lastSearchText && Asc.scope.lastPageNumber > 0) {
                  // 跳转到指定页面
                  oDocument.GoToPage(Asc.scope.lastPageNumber - 1);

                  // 搜索文本
                  const oSearchResults = oDocument.Search(Asc.scope.lastSearchText, false);
                  let clearedCount = 0;

                  // 遍历所有搜索结果，清除该页面的高亮
                  for (let i = 0; i < oSearchResults.length; i++) {
                    const oSearchResult = oSearchResults[i];
                    const page = oSearchResult.GetStartPage() + 1;

                    // 只清除指定页面的高亮
                    if (page === Asc.scope.lastPageNumber) {
                      oSearchResult.SetHighlight('none');
                      clearedCount++;
                    }
                  }

                  console.log('[Plugin] 已清除第', Asc.scope.lastPageNumber, '页的', clearedCount, '个高亮');

                  return {
                    success: true,
                    clearedCount: clearedCount,
                    page: Asc.scope.lastPageNumber
                  };
                } else {
                  // 如果没有指定搜索文本和页码，清空所有搜索（兼容旧逻辑）
                  oDocument.Search('', false);

                  console.log('[Plugin] 已清除所有高亮');

                  return {
                    success: true
                  };
                }
              } catch (error) {
                console.error('[Plugin] clearHighlights 执行错误:', error);
                return {
                  success: false,
                  error: error.message || '清除高亮失败'
                };
              }
            },
            false,
            true,
            function (result) {
              console.log('[Plugin] clearHighlights 执行完成, RequestId:', Asc.scope.clearRequestId, 'result:', result);
              if (result && result.success) {
                const message = result.clearedCount
                  ? `已清除第 ${result.page} 页的 ${result.clearedCount} 个高亮`
                  : '已清除所有高亮';
                Asc.scope.clearSendResult(Asc.scope.clearRequestId, 'success', { message: message }, null);
              } else {
                Asc.scope.clearSendResult(Asc.scope.clearRequestId, 'error', null, result?.error || '清除高亮失败');
              }
            }
          );
          break

        case 'toggleNavigation':
          // 切换导航面板显示/隐藏
          console.log('[Plugin] toggleNavigation 开始执行, RequestId:', requestId);

          window.Asc.plugin.executeMethod('ToggleLeftPanel', null, function (result) {
            console.log('[Plugin] toggleNavigation 执行完成, RequestId:', requestId, 'result:', result);
            if (result !== false) {
              sendResult(requestId, 'success', {
                message: '导航面板切换成功',
                panelVisible: result
              }, null);
            }
          });
          break

        default:
          console.warn(
            '[Plugin] 收到未知命令, RequestId:',
            requestId,
            'command:',
            command
          )
          sendResult(requestId, 'error', null, '未知命令: ' + command)
      }
    } catch (error) {
      console.error(
        '[Plugin] handleCommand 顶层异常, RequestId:',
        requestId,
        'error:',
        error
      )
      sendResult(requestId, 'error', null, error.message)
    }
  }

  // 初始化 WebSocket 连接
  function init() {
    // 等待插件环境就绪
    if (window.Asc && window.Asc.plugin) {
      console.log('[WebSocket] 插件环境已就绪，开始连接 WebSocket...')
      connect()
    } else {
      // 如果插件环境未就绪，延迟初始化（最多等待 10 秒）
      var retryCount = 0
      var maxRetries = 100 // 100 * 100ms = 10秒
      var checkInterval = setInterval(function () {
        retryCount++
        if (window.Asc && window.Asc.plugin) {
          clearInterval(checkInterval)
          console.log('[WebSocket] 插件环境已就绪，开始连接 WebSocket...')
          connect()
        } else if (retryCount >= maxRetries) {
          clearInterval(checkInterval)
          console.error('[WebSocket] 插件环境初始化超时，无法连接 WebSocket')
        }
      }, 100)
    }
  }

  // 页面加载完成后初始化（后台模式也适用）
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init)
  } else {
    // 如果文档已加载，立即初始化
    init()
    //
  }

  // 暴露给全局，方便调试
  window.WSClient = {
    connect: connect,
    sendResult: sendResult,
    getConnection: function () {
      return ws
    },
  }
})(window, undefined)

