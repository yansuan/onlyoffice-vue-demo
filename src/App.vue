<template>
  <div class="app">
    <div class="editor-container">
      <!-- <div class="mode-bar">
        <span>当前模式：{{ currentMode }}</span>
        <div class="mode-buttons">
          <button
            class="btn btn-small"
            :class="{ active: currentMode === 'edit' }"
            @click="switchMode('edit')"
          >
            编辑模式
          </button>
          <button
            class="btn btn-small"
            :class="{ active: currentMode === 'review' }"
            @click="switchMode('review')"
          >
            审阅模式
          </button>
          <button
            class="btn btn-small"
            :class="{ active: currentMode === 'view' }"
            @click="switchMode('view')"
          >
            只读模式
          </button>
        </div>
      </div> -->
      <DocumentEditor
        v-if="config"
        ref="docEditor"
        id="docEditor"
        documentServerUrl="http://192.168.93.128:8101"
        :config="config"
        :events_onAppReady="onAppReady"
        :events_onDocumentReady="onDocumentReady"
        :events_onRequestHistory="onRequestHistory"
        :events_onRequestHistoryData="onRequestHistoryData"
        :events_onRequestRestore="onRequestRestore"
        :events_onRequestHistoryClose="onRequestHistoryClose"
        :events_onRequestCompareFile="onRequestCompareFile"
        :onLoadComponentError="onLoadComponentError"
      />
      <div v-else class="loading">正在生成文档配置...</div>
    </div>
    <!-- 外部 JsAPI 执行器面板 -->
    <div class="jsapi-panel">
      <div class="panel-header">
        <h3>功能操作</h3>
        <!-- <div class="ws-status" :class="{ connected: wsConnected }">
          {{ wsConnected ? '已连接' : '未连接' }}
        </div> -->
      </div>
      <div class="panel-content">
        <!-- 快速操作
        <div class="section">
          <h4>快速操作</h4>
          <div class="input-group">
            <input v-model="quickText" type="text" placeholder="输入要插入的文本" class="input" />
            <button @click="handleInsertText" class="btn btn-primary">插入文本</button>
          </div>

          <div class="input-group">
            <input v-model="searchText" type="text" placeholder="输入要搜索的文本" class="input" />
            <button @click="handleSearchText" class="btn btn-primary">搜索文本</button>
          </div>
        </div>
 -->
        <!-- 批注操作 -->
        <div class="section">
          <h4>批注操作</h4>
          <button @click="handleGetAllComments" class="btn btn-primary">获取所有批注</button>
          <!-- <div class="input-group">
            <input v-model="commentText" type="text" placeholder="批注内容" class="input" />
            <button @click="handleAddComment" class="btn btn-primary">添加批注</button>
          </div> -->
          <div class="input-group" style="margin-top: 8px">
            <input
              v-model="selectionCommentText"
              type="text"
              placeholder="批注内容（在选中内容上添加）"
              class="input"
            />
            <button @click="handleAddCommentToSelection" class="btn btn-primary">
              在选中内容添加批注
            </button>
          </div>
        </div>

        <!-- 段落操作 -->
        <!-- <div class="section">
          <h4>段落操作</h4>
          <button @click="handleGetAllParagraphs" class="btn btn-secondary">获取所有段落</button>
        </div> -->

        <!-- 页面导航 -->
        <div class="section">
          <h4>高亮</h4>
          <div class="input-group">
            <input
              v-model.number="pageNumber"
              type="number"
              min="1"
              placeholder="输入页码"
              class="input"
              @keyup.enter="handleScrollToPage"
            />
            <!-- <button @click="handleScrollToPage" class="btn btn-primary">跳转到页</button> -->
          </div>
          <div class="input-group" style="margin-top: 8px">
            <input
              v-model="pageSearchText"
              type="text"
              placeholder="搜索内容并跳转"
              class="input"
              @keyup.enter="handleSearchAndNavigate"
            />
          </div>
          <div style="margin-top: 8px">
            <button @click="handleSearchAndNavigate" class="btn btn-primary">搜索并高亮</button>
            <button @click="handleClearHighlights" class="btn btn-warning">清除高亮</button>
          </div>
          <!-- 搜索结果显示 -->
          <div v-if="searchResults.length > 0" class="search-result-info">
            <span>在第 {{ pageNumber }} 页找到 {{ searchResults.length }} 个匹配项</span>
          </div>
        </div>

        <!-- 文档操作（已隐藏）
        <div class="section">
          <h4>文档操作</h4>
          <button @click="handleForceSave" class="btn btn-primary">手动保存 (forcesave)</button>
          <button @click="handleDownloadDocument" class="btn btn-secondary">
            下载文档到服务器
          </button>
        </div>
        -->

        <!-- 结果显示 -->
        <div class="section">
          <h4>执行结果</h4>
          <div class="result-box">
            <pre v-if="lastResult"
              >{{ JSON.stringify(lastResult, null, 2) }}
            </pre>
            <div v-else class="result-placeholder">暂无结果</div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script lang="ts">
import { defineComponent } from 'vue'
import { DocumentEditor } from '@onlyoffice/document-editor-vue'
import CryptoJS from 'crypto-js'

type CryptoWordArray = CryptoJS.lib.WordArray
type HistoryVersion = {
  created: string
  key: string
  user: {
    id: string
    name: string
  }
  version: number
  changes?: string // 变更描述，例如："修改了第3段，添加了5行"
}

type HistoryData = {
  currentVersion: number
  history: HistoryVersion[]
}

type EditorInstance = {
  refreshHistory?: (data: HistoryData) => void
  save?: () => void
  [key: string]: any
}

const ONLY_OFFICE_SECRET = '+keng2vx4V2ei1k/2wAsbxjpNP/v6Ew7uhyaJ9hgOr4='

const baseConfig = {
  document: {
    fileType: 'docx',
    // key 在 buildConfigWithToken 中动态生成，避免版本冲突提示
    title: 'OnlyOffice Vue Demo.docx',
    // 加载自动保存文件（demo_autosave.docx），自动保存会持续更新此文件
    url: 'http://192.168.93.1:4000/files/demo.docx',
    permissions: {
      edit: true,
      review: false,
      view: false,
      // 明确禁止自动保存相关权限
      download: true,
      print: true,
      copy: true,
      comment: true,
      fillForms: true,
      modifyFilter: true,
      modifyContentControl: true,
      changeHistory: true,
    },
  },
  documentType: 'word',
  editorConfig: {
    lang: 'zh',
    callbackUrl: 'http://192.168.93.1:4000/onlyoffice/callback',
    mode: 'edit',
    trackChanges: true,
    showReviewChanges: true,
    // 协同编辑配置
    coEditing: {
      mode: 'fast', // 快速模式，允许自动保存
      change: true, // 启用自动保存变更
    },
    customization: {
      autosave: true,
      compactHeader: true,
      compactToolbar: true,
      hideFileMenu: true, // 隐藏文件菜单
      rightMenu: false, // 隐藏右侧设置面板
      comments: true, // 只保留批注功能（可选，根据需求决定是否显示）
      feedback: false,
      help: false,
      about: false,
      features: {
        tabBackground: 'toolbar',
        tabStyle: 'line',
      },
      layout: {
        leftMenu: {
          mode: true, // 左侧面板默认展开
          navigation: true, // 显示“导航/标题”按钮
        },
      },
      css: `
          .tabs { display: none !important; }
          .header { display: none !important; }
          /* 默认打开左侧导航面板 */
          .asc-window.left-panel { display: block !important; }
          .asc-window-content.left-panel-open { margin-left: 300px !important; }
          .left-panel { width: 300px !important; display: block !important; }
          .left-panel .panel { display: block !important; }
          .left-panel .navigation { display: block !important; }
          .left-panel .headings { display: block !important; }
          /* 强制显示导航面板内容 */
          #left-panel-navigation { display: block !important; }
          .navigation-panel { display: block !important; }
          /* 其他样式 */
        `,
    },
    user: {
      id: 'user-1001',
      name: '测试用户',
      avatarUrl: 'https://cdn.jsdelivr.net/gh/baimingxuan/media-assets/avatar-default.png',
      permissions: {
        edit: true,
        comment: true,
        review: false,
      },
      // 自定义加：协作人角色，可选值如 'owner' | 'editor' | 'viewer' | 'reviewer'
      role: 'editor',
    },
    collaborators: [
      {
        id: 'user-1001',
        name: '测试用户',
        avatarUrl: 'https://cdn.jsdelivr.net/gh/baimingxuan/media-assets/avatar-default.png',
        role: 'reviewer',
      },
      {
        id: 'user-1002',
        name: '张三',
        avatarUrl: 'https://cdn.jsdelivr.net/gh/baimingxuan/media-assets/avatar-default.png',
        role: 'viewer',
      },
    ],
    plugins: {
      autostart: ['asc.{jsapi-executor-1234-5678-90ab-cdef12345678}'],
      visible: false,
    },
    review: {
      hideReviewDisplay: false,
      showReviewChanges: true,
      reviewDisplay: 'markup',
      trackChanges: true,
      hoverMode: false,
    },
  },
}

const base64Url = (word: CryptoWordArray | string) =>
  CryptoJS.enc.Base64.stringify(typeof word === 'string' ? CryptoJS.enc.Utf8.parse(word) : word)
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '')

const buildToken = (config: { document: unknown; documentType: string; editorConfig: unknown }) => {
  const header = { alg: 'HS256', typ: 'JWT' }
  const payload = {
    document: config.document,
    documentType: config.documentType,
    editorConfig: config.editorConfig,
    exp: Math.floor(Date.now() / 1000) + 30 * 60,
  }

  const headerEncoded = base64Url(JSON.stringify(header))
  const payloadEncoded = base64Url(JSON.stringify(payload))
  const signature = base64Url(
    CryptoJS.HmacSHA256(`${headerEncoded}.${payloadEncoded}`, ONLY_OFFICE_SECRET),
  )

  return `${headerEncoded}.${payloadEncoded}.${signature}`
}

const buildConfigWithToken = () => {
  // 使用时间戳生成动态 key，每次打开都是新的编辑会话
  const timestamp = Date.now()
  const documentKey = `vue-demo-doc-${timestamp}`
  // const documentKey = 'vue-demo-doc-demo-test'
  console.log('documentKey:', documentKey)

  // 在 URL 后面加时间戳参数，强制浏览器和 OnlyOffice 不使用缓存
  // const documentUrl = `${baseConfig.document.url}?t=${timestamp}`
  const documentUrl = `${baseConfig.document.url}`

  const document = {
    ...baseConfig.document,
    key: documentKey,
    url: documentUrl, // 使用带时间戳的 URL
  }

  const config = {
    ...baseConfig,
    document,
  }

  return {
    ...config,
    token: buildToken(config),
  }
}

type WSMessage = {
  requestId?: string // 请求 ID，用于匹配请求和响应
  type: 'success' | 'error' | 'connected'
  result?: unknown
  error?: string
  clientType?: string
  durationMs?: number // 调用耗时（毫秒），由发送端计算
}

export default defineComponent({
  name: 'App',
  components: { DocumentEditor },
  data() {
    return {
      config: null as null | ReturnType<typeof buildConfigWithToken>,
      currentMode: 'edit' as 'edit' | 'review' | 'view',
      editorInstance: null as EditorInstance | null,
      documentHistory: [] as HistoryVersion[],
      // WebSocket 相关
      ws: null as WebSocket | null,
      wsConnected: false,
      commandIdCounter: 0,
      pendingCommands: new Map<string, (result: WSMessage) => void>(),
      // UI 数据
      quickText: '',
      searchText: '对象代表查找操作的执行条件',
      commentText: '',
      selectionCommentText: '',
      pageNumber: 5,
      pageSearchText: '对象代表查找操作的执行条件',
      searchResults: [] as Array<{ text: string; page: number; index: number }>,
      lastResult: null as unknown | null,
    }
  },
  created() {
    this.config = buildConfigWithToken()
    if (
      this.config &&
      (this.config as any).editorConfig &&
      (this.config as any).editorConfig.mode
    ) {
      this.currentMode = (this.config as any).editorConfig.mode as 'edit' | 'review' | 'view'
    }
    this.initWebSocket()

    // 监听页面刷新/关闭事件
    window.addEventListener('beforeunload', this.handleBeforeUnload)
  },
  beforeUnmount() {
    console.log('[生命周期] beforeUnmount - 组件即将卸载')

    // 移除 beforeunload 事件监听
    window.removeEventListener('beforeunload', this.handleBeforeUnload)

    // 销毁编辑器和关闭连接
    this.cleanup()
  },
  methods: {
    switchMode(mode: 'edit' | 'review' | 'view') {
      if (this.currentMode === mode) {
        return
      }

      // 更新 baseConfig 的模式和相关权限
      // 注意：OnlyOffice 的 editorConfig.mode 只支持 'edit' 和 'view'
      // 这里的 'review' 是前端自定义的逻辑模式，底层仍然使用 'edit'
      if (mode === 'view') {
        baseConfig.editorConfig.mode = 'view'
      } else if (mode === 'review') {
        baseConfig.editorConfig.mode = 'review'
      } else {
        baseConfig.editorConfig.mode = 'edit'
      }

      if (mode === 'edit') {
        baseConfig.document.permissions.edit = true
        baseConfig.document.permissions.review = false
        baseConfig.editorConfig.user.permissions.edit = true
        baseConfig.editorConfig.user.permissions.review = false
        baseConfig.editorConfig.review.trackChanges = true
        baseConfig.editorConfig.review.showReviewChanges = true
      } else if (mode === 'review') {
        baseConfig.document.permissions.edit = false
        baseConfig.document.permissions.review = true
        baseConfig.editorConfig.user.permissions.edit = true
        baseConfig.editorConfig.user.permissions.review = true
        baseConfig.editorConfig.review.trackChanges = true
        baseConfig.editorConfig.review.showReviewChanges = true
      } else if (mode === 'view') {
        baseConfig.document.permissions.edit = false
        baseConfig.document.permissions.review = false
        baseConfig.editorConfig.user.permissions.edit = false
        baseConfig.editorConfig.user.permissions.review = false
        baseConfig.editorConfig.review.trackChanges = true
        baseConfig.editorConfig.review.showReviewChanges = true
      }

      // 重新生成配置和 token
      this.config = buildConfigWithToken()
      this.currentMode = mode
      console.log('[OnlyOffice] 模式已切换为:', mode)
    },
    onAppReady(editorInstance: EditorInstance | null) {
      console.log('onAppReady', editorInstance)
      this.editorInstance = editorInstance
    },
    onDocumentReady() {
      console.log('Document is loaded')

      // 文档加载完成后，自动打开导航面板（标题列表）
      setTimeout(() => {
        this.openNavigationPanel()
      }, 1000) // 延迟1秒确保编辑器完全加载
    },
    onLoadComponentError(errorCode: number, errorDescription: string) {
      console.error(`Editor load error ${errorCode}: ${errorDescription}`)
    },
    // 历史记录相关方法
    onRequestHistory() {
      console.log('用户请求查看历史记录')
      // 这里应该从服务器获取历史记录，demo 使用模拟数据
      const historyData: HistoryData = {
        currentVersion: this.documentHistory.length || 1,
        history: this.documentHistory.length
          ? this.documentHistory
          : [
              {
                created: new Date(Date.now() - 86400000)
                  .toISOString()
                  .replace('T', ' ')
                  .slice(0, 19),
                key: 'version1',
                user: {
                  id: 'user1',
                  name: 'User 1',
                },
                version: 1,
                changes: '初始版本',
              },
              {
                created: new Date(Date.now() - 3600000)
                  .toISOString()
                  .replace('T', ' ')
                  .slice(0, 19),
                key: 'version2',
                user: {
                  id: 'user2',
                  name: 'User 2',
                },
                version: 2,
                changes: '修改了第2段，添加了3行文本；删除了第5段的部分内容',
              },
            ],
      }

      // 调用 refreshHistory 显示历史记录面板
      if (this.editorInstance?.refreshHistory) {
        this.editorInstance.refreshHistory(historyData)
      } else {
        // 如果 refreshHistory 不存在，尝试从 window.DocEditor 获取
        const editor = (window as { DocEditor?: { instances?: { [key: string]: EditorInstance } } })
          .DocEditor?.instances?.['docEditor']
        if (editor?.refreshHistory) {
          editor.refreshHistory(historyData)
        } else {
          console.warn('refreshHistory method not available')
        }
      }
    },
    onRequestHistoryData(event: { data: { version: number; key: string } }) {
      console.log('用户请求查看历史版本数据:', event.data)
      const { version, key } = event.data

      // 返回该版本的文件信息
      // 实际应用中应该从服务器获取该版本的文件 URL
      return {
        fileType: 'docx',
        key: key,
        url: `https://example-files.online-convert.com/document/docx/example.docx?version=${version}`,
        version: version,
      }
    },
    onRequestRestore(event: { data: { version: number; key: string } }) {
      console.log('用户请求恢复版本:', event.data)
      const { version, key } = event.data

      // 这里应该处理版本恢复逻辑
      // 1. 从服务器获取该版本的文件
      // 2. 替换当前文档
      // 3. 保存为新版本

      alert(
        `恢复版本 ${version} (key: ${key})\n实际应用中需要从服务器获取该版本文件并替换当前文档。`,
      )
    },
    onRequestHistoryClose() {
      console.log('用户关闭历史记录面板')
    },
    // 文档比较功能 - 用于显示两个版本之间的差异
    onRequestCompareFile(event: {
      data: {
        fileType: string
        url: string
        key: string
        version: number
      }
    }) {
      console.log('用户请求比较文档版本:', event.data)
      const { fileType, url, key, version } = event.data

      // 返回要比较的文件信息
      // 这里应该返回历史版本的文件 URL，OnlyOffice 会自动高亮显示差异
      return {
        fileType: fileType,
        key: key,
        url: url, // 历史版本的文件 URL
        version: version,
      }
    },
    // WebSocket 相关方法
    initWebSocket() {
      const wsUrl = 'ws://192.168.93.1:4000?type=vue'
      try {
        this.ws = new WebSocket(wsUrl)

        this.ws.onopen = () => {
          console.log('[WebSocket] Vue 应用已连接到服务器')
          this.wsConnected = true
        }

        this.ws.onmessage = (event) => {
          try {
            const data = JSON.parse(event.data) as WSMessage
            console.log('[WebSocket] 收到结果:', data)

            // 处理连接确认消息
            if (data.type === 'connected') {
              return
            }

            // 处理命令结果（带超时与一次性回调控制）
            const requestId = data.requestId
            if (requestId) {
              if (this.pendingCommands.has(requestId)) {
                const callback = this.pendingCommands.get(requestId)
                if (callback) {
                  callback(data)
                  this.pendingCommands.delete(requestId)
                }
              } else {
                // 该 requestId 已经超时或被处理过，后续结果全部忽略（不再更新 UI、不再回调）
                console.warn('[WebSocket] 收到已过期或已处理的 RequestId 响应，忽略:', requestId)
                return
              }
            }

            // 仅对未绑定 requestId 的广播类消息，或正常处理中的结果，更新最后结果
            this.lastResult = data
          } catch (error) {
            console.error('[WebSocket] 消息解析错误:', error)
          }
        }

        this.ws.onerror = (error) => {
          console.error('[WebSocket] 连接错误:', error)
          this.wsConnected = false
        }

        this.ws.onclose = () => {
          console.log('[WebSocket] 连接已关闭，尝试重连...')
          this.wsConnected = false
          // 3秒后重连
          setTimeout(() => {
            this.initWebSocket()
          }, 3000)
        }
      } catch (error) {
        console.error('[WebSocket] 初始化失败:', error)
        this.wsConnected = false
      }
    },
    sendCommand(command: string, params: Record<string, unknown> = {}): Promise<WSMessage> {
      return new Promise((resolve, reject) => {
        if (!this.ws || this.ws.readyState !== WebSocket.OPEN) {
          reject(new Error('【前端】WebSocket 未连接'))
          return
        }

        const startTime = Date.now()

        // 生成唯一的 RequestId
        const requestId = `req-${++this.commandIdCounter}-${Date.now()}`
        const message = {
          requestId,
          command,
          params,
        }

        // 存储回调
        this.pendingCommands.set(requestId, (result) => {
          if (result.type === 'error') {
            const msg = result.error ? `【插件】${result.error}` : '【插件】执行失败'
            reject(new Error(msg))
          } else {
            resolve(result)
          }
        })

        // 发送命令
        this.ws.send(JSON.stringify(message))
        console.log('[WebSocket] 发送命令 RequestId:', requestId, '命令:', command)

        // 30秒超时处理
        const timeoutId = setTimeout(() => {
          if (this.pendingCommands.has(requestId)) {
            this.pendingCommands.delete(requestId)
            const timeoutError = new Error(`【前端】命令执行超时（超过 30 秒）: ${command}`)
            console.error('[WebSocket] 命令超时:', requestId, command)
            reject(timeoutError)
          }
        }, 30000)

        // 如果命令完成，清除超时定时器
        const originalCallback = this.pendingCommands.get(requestId)
        if (originalCallback) {
          this.pendingCommands.set(requestId, (result) => {
            clearTimeout(timeoutId)
            const durationMs = Date.now() - startTime
            const resultWithDuration: WSMessage = {
              ...result,
              durationMs,
            }
            originalCallback(resultWithDuration)
          })
        }
      })
    },
    // UI 操作方法
    async handleInsertText() {
      if (!this.quickText.trim()) {
        alert('请输入要插入的文本')
        return
      }

      try {
        const result = await this.sendCommand('insertText', { text: this.quickText })
        console.log('插入文本成功 RequestId:', result.requestId, result)
        this.quickText = ''
      } catch (error) {
        console.error('插入文本失败:', error)
        const errorMsg = (error as Error).message
        alert('插入文本失败: ' + errorMsg)
      }
    },
    async handleSearchText() {
      if (!this.searchText.trim()) {
        alert('请输入要搜索的文本')
        return
      }

      try {
        const result = await this.sendCommand('searchText', { text: this.searchText })
        console.log('搜索文本成功 RequestId:', result.requestId, result)
      } catch (error) {
        console.error('搜索文本失败:', error)
        const errorMsg = (error as Error).message
        alert('搜索文本失败: ' + errorMsg)
      }
    },
    async handleGetAllComments() {
      try {
        const result = await this.sendCommand('getAllComments')
        console.log('获取批注成功 RequestId:', result.requestId, result)
      } catch (error) {
        console.error('获取批注失败:', error)
        const errorMsg = (error as Error).message
        alert('获取批注失败: ' + errorMsg)
      }
    },
    async handleAddComment() {
      if (!this.commentText.trim()) {
        alert('请输入批注内容')
        return
      }

      try {
        const result = await this.sendCommand('addComment', {
          userName: 'Vue User',
          quoteText: this.commentText,
          text: this.commentText,
        })
        console.log('添加批注成功 RequestId:', result.requestId, result)
        this.commentText = ''
      } catch (error) {
        console.error('添加批注失败:', error)
        const errorMsg = (error as Error).message
        alert('添加批注失败: ' + errorMsg)
      }
    },
    async handleAddCommentToSelection() {
      if (!this.selectionCommentText.trim()) {
        alert('请输入批注内容')
        return
      }

      try {
        const result = await this.sendCommand('addCommentToSelection', {
          text: this.selectionCommentText,
          userName: 'Vue User',
        })
        console.log('在选中内容添加批注成功 RequestId:', result.requestId, result)

        // 显示成功信息，包含选中的文本
        const resultData = result.result as { selectedText?: string } | null
        if (resultData && resultData.selectedText) {
          alert(`批注已添加到选中内容："${resultData.selectedText}"`)
        } else {
          alert('批注添加成功')
        }

        this.selectionCommentText = ''
      } catch (error) {
        console.error('在选中内容添加批注失败:', error)
        const errorMsg = (error as Error).message
        alert('添加批注失败: ' + errorMsg)
      }
    },
    // 获取所有段落
    handleGetAllParagraphs() {
      this.sendCommand('getAllParagraphs')
        .then((result) => {
          this.lastResult = result
        })
        .catch((error) => {
          console.error('获取段落失败:', error)
          this.lastResult = { error: '获取段落失败', details: error }
        })
    },
    // 跳转到指定页码
    handleScrollToPage() {
      if (!this.pageNumber || this.pageNumber < 1) {
        this.lastResult = { error: '请输入有效的页码' }
        return
      }

      this.sendCommand('scrollToPage', { pageNumber: this.pageNumber })
        .then((result) => {
          this.lastResult = result
        })
        .catch((error) => {
          console.error('跳转页面失败:', error)
          this.lastResult = { error: '跳转页面失败', details: error }
        })
    },
    // 搜索并跳转到页面
    async handleSearchAndNavigate() {
      if (!this.pageSearchText.trim()) {
        alert('请输入要搜索的内容')
        return
      }

      if (!this.pageNumber || this.pageNumber < 1) {
        alert('请输入有效的页码（必填）')
        return
      }

      try {
        const result = await this.sendCommand('searchAndNavigate', {
          searchText: this.pageSearchText,
          pageNumber: this.pageNumber,
        })

        if (result.result && result.result.results && result.result.results.length > 0) {
          this.searchResults = result.result.results

          this.lastResult = {
            success: true,
            message:
              result.result.message ||
              `在第 ${this.pageNumber} 页找到 ${result.result.count} 个匹配项`,
            results: result.result.results,
            count: result.result.count,
          }
        } else {
          this.searchResults = []
          this.lastResult = { message: '未找到匹配的内容' }
          alert('未找到匹配的内容')
        }
      } catch (error) {
        console.error('搜索失败:', error)
        const errorMsg = (error as Error).message
        alert(errorMsg)
        this.lastResult = { error: '搜索失败', details: errorMsg }
        this.searchResults = []
      }
    },
    // 清除高亮
    async handleClearHighlights() {
      try {
        // 传递当前搜索的文本和页码，以便精确清除高亮
        const result = await this.sendCommand('clearHighlights', {
          searchText: this.pageSearchText,
          pageNumber: this.pageNumber,
        })

        this.searchResults = []

        if (result.result && result.result.message) {
          this.lastResult = { success: true, message: result.result.message }
        } else {
          this.lastResult = { success: true, message: '已清除所有高亮' }
        }
      } catch (error) {
        console.error('清除高亮失败:', error)
        const errorMsg = (error as Error).message
        alert('清除高亮失败: ' + errorMsg)
      }
    },
    async handleForceSave() {
      try {
        console.log('[手动保存] 开始触发 forcesave')

        // 获取当前编辑器正在使用的 document key
        const documentKey = this.config?.document.key
        if (!documentKey) {
          throw new Error('无法获取 document key')
        }

        console.log('[手动保存] 文档 key:', documentKey)

        // 通过后端 API 调用 forcesave（避免 CORS 问题）
        const response = await fetch('http://192.168.93.1:4000/api/forcesave', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            documentKey: documentKey,
          }),
        })

        const result = await response.json()
        console.log('[手动保存] API 响应:', result)

        if (result.success) {
          alert(
            '✓ 手动保存已触发\n\n' +
              `Document Key: ${documentKey}\n` +
              'OnlyOffice 将通过 callback 保存文档\n\n' +
              '请查看 callback-server 控制台日志',
          )
        } else {
          alert('保存失败: ' + (result.error || '未知错误'))
        }
      } catch (error) {
        console.error('手动保存失败:', error)
        const errorMsg = (error as Error).message
        alert('手动保存失败: ' + errorMsg)
      }
    },
    async handleDownloadDocument() {
      try {
        console.log('[下载文档] 开始下载正在编辑的文档')

        // 获取当前编辑器正在使用的 document key
        const documentKey = this.config?.document.key
        if (!documentKey) {
          throw new Error('无法获取 document key')
        }

        console.log('[下载文档] 文档 key:', documentKey)

        // 调用插件的 downloadDocument 命令
        const result = await this.sendCommand('downloadDocument', { documentKey })
        console.log('[下载文档] 插件返回结果:', result)

        if (result.type === 'success' && result.result?.success) {
          const filePath = result.result.filePath || 'downloaded_document.docx'
          const fileSize = result.result.fileSize
            ? (result.result.fileSize / 1024).toFixed(2) + ' KB'
            : '未知'
          alert(
            `✓ 文档已成功下载到服务器\n\n` +
              `文件路径: ${filePath}\n` +
              `文件大小: ${fileSize}\n\n` +
              `请查看 callback-server 控制台日志`,
          )
        } else {
          alert('下载失败: ' + (result.error || result.result?.message || '未知错误'))
        }
      } catch (error) {
        console.error('下载文档失败:', error)
        const errorMsg = (error as Error).message
        alert('下载文档失败: ' + errorMsg)
      }
    },
    // 页面刷新/关闭前的处理
    handleBeforeUnload() {
      console.log('[页面事件] beforeunload - 页面即将刷新/关闭')

      // 执行清理操作
      this.cleanup()
    },
    // 清理资源
    cleanup() {
      console.log('[清理] 开始清理资源...')

      // 1. 销毁编辑器实例
      if (this.editorInstance) {
        try {
          console.log('[清理] 正在销毁编辑器实例...')
          this.editorInstance.destroyEditor()
          this.editorInstance = null
          console.log('[清理] ✓ 编辑器实例已销毁')
        } catch (error) {
          console.error('[清理] 销毁编辑器失败:', error)
        }
      }

      // 2. 关闭 WebSocket 连接
      if (this.ws) {
        try {
          console.log('[清理] 正在关闭 WebSocket 连接...')
          this.ws.close()
          this.ws = null
          console.log('[清理] ✓ WebSocket 连接已关闭')
        } catch (error) {
          console.error('[清理] 关闭 WebSocket 失败:', error)
        }
      }

      console.log('[清理] ✓ 资源清理完成')
    },
  },
})
</script>

<style scoped>
.app {
  width: 100vw;
  height: 100vh;
  margin: 0;
  display: flex;
  flex-direction: row;
}

.editor-container {
  flex: 1 1 auto;
  display: flex;
  flex-direction: column;
  min-width: 0;
}

.loading {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  height: 100%;
  font-size: 16px;
  color: #555;
  letter-spacing: 0.5px;
}

#docEditor {
  flex: 1 1 auto;
  width: 100%;
  border: 0;
}

/* JsAPI 执行器面板 */
.jsapi-panel {
  width: 400px;
  background: #f5f7fb;
  border-left: 1px solid #e0e4ee;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.panel-header {
  padding: 16px;
  background: #fff;
  border-bottom: 1px solid #e0e4ee;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.panel-header h3 {
  margin: 0;
  font-size: 16px;
  font-weight: 600;
  color: #333;
}

.ws-status {
  padding: 4px 12px;
  border-radius: 12px;
  font-size: 12px;
  background: #dc3545;
  color: #fff;
}

.ws-status.connected {
  background: #28a745;
}

.panel-content {
  flex: 1 1 auto;
  overflow-y: auto;
  padding: 16px;
}

.section {
  margin-bottom: 24px;
}

.section h4 {
  margin: 0 0 12px 0;
  font-size: 14px;
  font-weight: 600;
  color: #555;
}

.input-group {
  display: flex;
  gap: 8px;
  margin-bottom: 8px;
}

.input {
  flex: 1 1 auto;
  padding: 8px 12px;
  border: 1px solid #e0e4ee;
  border-radius: 4px;
  font-size: 13px;
}

.btn {
  padding: 8px 16px;
  border: none;
  border-radius: 4px;
  font-size: 13px;
  cursor: pointer;
  white-space: nowrap;
  margin-left: 8px;
  margin-right: 8px;
}

.btn-primary {
  background: #28a745;
  color: #fff;
}

.btn-primary:hover {
  background: #218838;
}

.btn-secondary {
  background: #6c757d;
}

.btn-secondary:hover {
  background: #5a6268;
}

.btn-warning {
  background: #ffc107;
  color: #000;
}

.btn-warning:hover {
  background: #e0a800;
}

.result-box {
  background: #fff;
  border: 1px solid #e0e4ee;
  border-radius: 4px;
  padding: 12px;
  max-height: 300px;
  overflow-y: auto;
}

.result-box pre {
  margin: 0;
  font-size: 12px;
  color: #333;
  white-space: pre-wrap;
  word-wrap: break-word;
}

.result-placeholder {
  color: #999;
  font-size: 12px;
  text-align: center;
  padding: 20px;
}

.search-result-info {
  margin-top: 10px;
  padding: 8px 12px;
  background-color: #e8f4f8;
  border-radius: 4px;
  font-size: 13px;
  color: #333;
  text-align: center;
}
</style>
