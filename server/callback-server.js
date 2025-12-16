import express from 'express'
import path from 'path'
import fs from 'fs'
import crypto from 'crypto'
import { fileURLToPath } from 'url'
import { createServer } from 'http'
import { WebSocketServer } from 'ws'

const app = express()
const PORT = process.env.ONLYOFFICE_CALLBACK_PORT || 4000

// 计算当前文件所在目录，方便配置静态文件目录
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// 静态文件目录：将 demo.docx 放在 server/files/ 目录下
const FILES_DIR = path.resolve(__dirname, 'files')
const ORIGINAL_FILE = path.join(FILES_DIR, 'demo.docx')  // 原始文件，不会被修改
const AUTOSAVE_FILE = path.join(FILES_DIR, 'demo_autosave.docx')  // 自动保存文件
const LATEST_FILE = path.join(FILES_DIR, 'demo_latest.docx')  // 手动保存的最新版本

// 初始化：如果 demo_autosave.docx 不存在，从 demo.docx 复制一份
if (!fs.existsSync(AUTOSAVE_FILE) && fs.existsSync(ORIGINAL_FILE)) {
  fs.copyFileSync(ORIGINAL_FILE, AUTOSAVE_FILE)
  console.log('[初始化] 从原始文件复制到自动保存文件:', AUTOSAVE_FILE)
}

// 初始化：如果 demo_latest.docx 不存在，从 demo.docx 复制一份
if (!fs.existsSync(LATEST_FILE) && fs.existsSync(ORIGINAL_FILE)) {
  fs.copyFileSync(ORIGINAL_FILE, LATEST_FILE)
  console.log('[初始化] 从原始文件复制到最新文件:', LATEST_FILE)
}

app.use(express.json({ limit: '100mb' }))
app.use(express.urlencoded({ extended: true }))

// 提供 demo.docx 的 HTTP 访问地址，例如：
// http://localhost:4000/files/demo.docx
app.use('/files', express.static(FILES_DIR))

// API: 接收文档内容并保存为文本文件
app.post('/api/save-content', async (req, res) => {
  try {
    const { content, timestamp } = req.body

    console.log('[API] 收到文档内容保存请求')
    console.log('[API] 时间戳:', timestamp)
    console.log('[API] 内容长度:', content ? content.length : 0, '字符')
    console.log('[API] 内容预览:', content ? content.substring(0, 100) + '...' : '空')

    // 保存为文本文件
    const textFilePath = path.join(FILES_DIR, 'document_content.txt')
    await fs.promises.writeFile(textFilePath, content, 'utf-8')

    const stats = await fs.promises.stat(textFilePath)
    console.log('[API] ✓ 文档内容已保存到:', textFilePath)
    console.log('[API] ✓ 文件大小:', (stats.size / 1024).toFixed(2), 'KB')

    res.json({
      success: true,
      message: '文档内容已保存',
      filePath: textFilePath,
      size: stats.size,
      timestamp: timestamp
    })
  } catch (error) {
    console.error('[API] 保存文档内容失败:', error)
    res.status(500).json({
      success: false,
      error: error.message
    })
  }
})

// API: 根据 document key 直接读取 OnlyOffice 缓存文件（最快方式）
app.post('/api/get-document-by-key', async (req, res) => {
  try {
    const { documentKey } = req.body

    console.log('[API] 收到根据 key 获取文档请求')
    console.log('[API] Document Key:', documentKey)

    // OnlyOffice 缓存目录（根据实际安装路径调整）
    // Docker: /var/lib/onlyoffice/documentserver/App_Data/cache/files
    // Windows: C:\Program Files\ONLYOFFICE\DocumentServer\App_Data\cache\files
    const ONLYOFFICE_CACHE_BASE = 'C:\\Program Files\\ONLYOFFICE\\DocumentServer\\App_Data\\cache\\files'

    // 计算 key 的 hash（OnlyOffice 使用 SHA-256）
    const keyHash = crypto.createHash('sha256').update(documentKey).digest('hex')
    console.log('[API] Key Hash:', keyHash)

    // 尝试多个可能的缓存路径
    const possiblePaths = [
      path.join(ONLYOFFICE_CACHE_BASE, keyHash, 'output.docx'),
      path.join(ONLYOFFICE_CACHE_BASE, keyHash, 'Editor.bin'),
      path.join(ONLYOFFICE_CACHE_BASE, keyHash.substring(0, 2), keyHash, 'output.docx'),
    ]

    let buffer = null
    let foundPath = null

    for (const cachePath of possiblePaths) {
      if (fs.existsSync(cachePath)) {
        console.log('[API] ✓ 找到缓存文件:', cachePath)
        buffer = await fs.promises.readFile(cachePath)
        foundPath = cachePath
        break
      }
    }

    if (buffer) {
      // 保存到临时文件
      const tempFilePath = path.join(FILES_DIR, 'temp_document.docx')
      await fs.promises.writeFile(tempFilePath, buffer)

      const stats = await fs.promises.stat(tempFilePath)
      console.log('[API] ✓ 文档已保存到:', tempFilePath)
      console.log('[API] ✓ 文件大小:', (stats.size / 1024).toFixed(2), 'KB')

      res.json({
        success: true,
        message: '文档已获取',
        filePath: tempFilePath,
        size: stats.size,
        cachePath: foundPath
      })
    } else {
      console.log('[API] ✗ 未找到缓存文件，尝试的路径:')
      possiblePaths.forEach(p => console.log('  -', p))
      res.status(404).json({
        success: false,
        error: '未找到文档缓存文件',
        triedPaths: possiblePaths
      })
    }
  } catch (error) {
    console.error('[API] 获取文档失败:', error)
    res.status(500).json({
      success: false,
      error: error.message
    })
  }
})

// API: 手动触发 forcesave
app.post('/api/forcesave', async (req, res) => {
  try {
    const { documentKey } = req.body

    console.log('[API] 收到 forcesave 请求')
    console.log('[API] Document Key:', documentKey)

    if (!documentKey) {
      return res.status(400).json({
        success: false,
        error: '缺少 documentKey 参数'
      })
    }

    // 调用 OnlyOffice Command Service API
    const commandServiceUrl = 'http://192.168.93.128:8101/coauthoring/CommandService.ashx'

    const commandPayload = {
      c: 'forcesave',
      key: documentKey,
      userdata: 'manual-forcesave'
    }

    console.log('[API] 调用 Command Service forcesave')
    console.log('[API] Payload:', commandPayload)

    const response = await fetch(commandServiceUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(commandPayload)
    })

    const result = await response.json()
    console.log('[API] Command Service 响应:', result)

    if (result.error === 0) {
      console.log('[API] ✓ forcesave 成功触发')
      res.json({
        success: true,
        message: 'forcesave 已触发',
        documentKey: documentKey,
        commandServiceResponse: result
      })
    } else {
      console.error('[API] ✗ forcesave 失败, error:', result.error)
      res.status(500).json({
        success: false,
        error: 'Command Service 返回错误: ' + result.error,
        commandServiceResponse: result
      })
    }
  } catch (error) {
    console.error('[API] forcesave 请求失败:', error)
    res.status(500).json({
      success: false,
      error: error.message
    })
  }
})

// API: 下载正在编辑的文档并保存到服务器
app.post('/api/download-document', async (req, res) => {
  try {
    const { fileData, fileName, timestamp } = req.body

    console.log('[API] 收到下载文档请求')
    console.log('[API] 文件名:', fileName || 'downloaded_document.docx')
    console.log('[API] 时间戳:', timestamp)
    console.log('[API] 数据长度:', fileData ? fileData.length : 0, '字符 (base64)')

    // 将 base64 数据转换为 Buffer
    const buffer = Buffer.from(fileData, 'base64')
    console.log('[API] 二进制大小:', (buffer.length / 1024).toFixed(2), 'KB')

    // 生成带时间戳的文件名
    const timestamp_str = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5)
    const downloadFileName = `downloaded_${timestamp_str}.docx`
    const downloadFilePath = path.join(FILES_DIR, downloadFileName)

    // 保存文件
    await fs.promises.writeFile(downloadFilePath, buffer)

    const stats = await fs.promises.stat(downloadFilePath)
    console.log('[API] ✓ 文档已下载并保存到:', downloadFilePath)
    console.log('[API] ✓ 文件大小:', (stats.size / 1024).toFixed(2), 'KB')
    console.log('[API] ✓ 保存时间:', stats.mtime.toISOString())

    res.json({
      success: true,
      message: '文档已成功下载到服务器',
      filePath: downloadFilePath,
      fileName: downloadFileName,
      fileSize: stats.size,
      timestamp: timestamp
    })
  } catch (error) {
    console.error('[API] 下载文档失败:', error)
    res.status(500).json({
      success: false,
      error: error.message
    })
  }
})

// 处理 ONLYOFFICE 回调：根据回调中的 url 下载并保存文档
app.post('/onlyoffice/callback', async (req, res) => {
  const body = req.body || {}
  console.log('[ONLYOFFICE] Callback payload:', JSON.stringify(body, null, 2))

  // 记录变更信息（如果存在）
  if (body.changes) {
    console.log('[ONLYOFFICE] Document changes detected:')
    console.log('  - Changes URL:', body.changes)
    console.log('  - Key:', body.key)
    console.log('  - Status:', body.status)
    // 实际应用中，应该下载 changes 文件并解析变更内容
    // 然后存储到数据库，用于在历史记录中显示修改位置
  }

  try {
    // ONLYOFFICE 文档状态：
    // status=1: 正在编辑
    // status=2: 文档已准备好保存（自动保存或关闭时）
    // status=3: 保存出错
    // status=4: 文档关闭且无更改
    // status=6: 正在保存（强制保存，通常是用户点击保存按钮）
    // status=7: 强制保存出错
    const status = body.status
    const fileUrl = body.url
    const key = body.key
    const actions = body.actions // 用户操作类型数组
    const forcesavetype = body.forcesavetype // 强制保存类型：0=定时, 1=用户点击保存, 2=命令服务, 3=提交表单

    console.log('[ONLYOFFICE] 回调详情:')
    console.log('  - status:', status)
    console.log('  - key:', key)
    console.log('  - url:', fileUrl ? '存在' : '不存在')
    console.log('  - forcesavetype:', forcesavetype)
    console.log('  - actions:', actions)

    // 处理自动保存 (status=2) 和手动保存 (status=6)
    const isAutoSave = status === 2 && fileUrl
    const isManualSave = status === 6 && fileUrl
    const hasUserAction = actions && actions.length > 0

    if (isAutoSave) {
      // 自动保存：保存到 demo_autosave.docx
      console.log('[ONLYOFFICE] ✓ 检测到自动保存回调 (status=2)')
      console.log('[ONLYOFFICE] 文档 key:', key)
      console.log('[ONLYOFFICE] 开始下载自动保存文档:', fileUrl)

      const response = await fetch(fileUrl)
      if (!response.ok) {
        throw new Error(`下载文档失败，HTTP 状态码: ${response.status}`)
      }

      const arrayBuffer = await response.arrayBuffer()
      const buffer = Buffer.from(arrayBuffer)

      const targetPath = AUTOSAVE_FILE
      console.log('[ONLYOFFICE] 准备写入自动保存文件到:', targetPath)
      console.log('[ONLYOFFICE] 文件大小:', buffer.length, 'bytes')

      await fs.promises.mkdir(FILES_DIR, { recursive: true })
      await fs.promises.writeFile(targetPath, buffer)

      const stats = await fs.promises.stat(targetPath)
      const fileSizeKB = (stats.size / 1024).toFixed(2)
      console.log('[ONLYOFFICE] ✓ 文档已自动保存到:', targetPath)
      console.log('[ONLYOFFICE] ✓ 文件大小:', fileSizeKB, 'KB')
      console.log('[ONLYOFFICE] ✓ 最后修改时间:', stats.mtime.toISOString())
    } else if (isManualSave) {
      // 手动保存：将 demo_autosave.docx 复制到 demo_latest.docx
      console.log('[ONLYOFFICE] ✓ 检测到手动保存请求 (status=6)')
      console.log('[ONLYOFFICE] 保存类型:', forcesavetype === 1 ? '用户点击保存按钮' : forcesavetype === 0 ? '定时强制保存' : forcesavetype === 2 ? '命令服务' : '其他')
      console.log('[ONLYOFFICE] userdata:', body.userdata)

      // 检查 userdata 判断操作类型
      if (body.userdata === 'manual-forcesave') {
        // 手动 forcesave 按钮触发
        console.log('[ONLYOFFICE] ✓ 确认为手动 forcesave 触发')
        console.log('[ONLYOFFICE] 文档 key:', key)

        // 先下载最新文档到 autosave 文件
        console.log('[ONLYOFFICE] 开始下载最新文档:', fileUrl)
        const response = await fetch(fileUrl)
        if (!response.ok) {
          throw new Error(`下载文档失败，HTTP 状态码: ${response.status}`)
        }

        const arrayBuffer = await response.arrayBuffer()
        const buffer = Buffer.from(arrayBuffer)

        await fs.promises.mkdir(FILES_DIR, { recursive: true })
        await fs.promises.writeFile(AUTOSAVE_FILE, buffer)
        console.log('[ONLYOFFICE] ✓ 已更新自动保存文件:', AUTOSAVE_FILE)

        // 复制 autosave 文件到 latest 文件
        fs.copyFileSync(AUTOSAVE_FILE, LATEST_FILE)

        const stats = await fs.promises.stat(LATEST_FILE)
        const fileSizeKB = (stats.size / 1024).toFixed(2)
        console.log('[ONLYOFFICE] ✓ 已将自动保存文件复制到最新文件:', LATEST_FILE)
        console.log('[ONLYOFFICE] ✓ 文件大小:', fileSizeKB, 'KB')
        console.log('[ONLYOFFICE] ✓ 最后修改时间:', stats.mtime.toISOString())
      } else if (body.userdata === 'temp-save') {
        // 暂存按钮触发
        console.log('[ONLYOFFICE] ✓ 确认为暂存按钮触发')

        // 下载文档并保存到临时文件
        console.log('[ONLYOFFICE] 开始下载文档内容:', fileUrl)
        const response = await fetch(fileUrl)
        if (!response.ok) {
          throw new Error(`下载文档失败，HTTP 状态码: ${response.status}`)
        }

        const arrayBuffer = await response.arrayBuffer()
        const buffer = Buffer.from(arrayBuffer)

        // 保存到临时文件
        const tempFilePath = path.join(FILES_DIR, 'temp_document.docx')
        await fs.promises.mkdir(FILES_DIR, { recursive: true })
        await fs.promises.writeFile(tempFilePath, buffer)

        const stats = await fs.promises.stat(tempFilePath)
        const fileSizeKB = (stats.size / 1024).toFixed(2)
        console.log('[ONLYOFFICE] ✓ 文档已暂存到:', tempFilePath)
        console.log('[ONLYOFFICE] ✓ 文件大小:', fileSizeKB, 'KB')
        console.log('[ONLYOFFICE] ✓ 最后修改时间:', stats.mtime.toISOString())
      } else if (body.userdata === 'download-document') {
        // 下载文档按钮触发
        console.log('[ONLYOFFICE] ✓ 确认为下载文档按钮触发')

        // 下载最新文档内容
        console.log('[ONLYOFFICE] 开始下载文档内容:', fileUrl)
        const response = await fetch(fileUrl)
        if (!response.ok) {
          throw new Error(`下载文档失败，HTTP 状态码: ${response.status}`)
        }

        const arrayBuffer = await response.arrayBuffer()
        const buffer = Buffer.from(arrayBuffer)

        // 保存到最终文件
        const finalFilePath = path.join(FILES_DIR, 'final_document.docx')
        await fs.promises.mkdir(FILES_DIR, { recursive: true })
        await fs.promises.writeFile(finalFilePath, buffer)

        const stats = await fs.promises.stat(finalFilePath)
        const fileSizeKB = (stats.size / 1024).toFixed(2)
        console.log('[ONLYOFFICE] ✓ 文档已保存到:', finalFilePath)
        console.log('[ONLYOFFICE] ✓ 文件大小:', fileSizeKB, 'KB')
        console.log('[ONLYOFFICE] ✓ 最后修改时间:', stats.mtime.toISOString())
      } else {
        console.log('[ONLYOFFICE] ⚠ 非手动操作触发的 forcesave，已忽略')
        console.log('[ONLYOFFICE] userdata:', body.userdata)
      }
    } else {
      console.log('[ONLYOFFICE] 回调状态无需保存文件，status:', status, 'url:', fileUrl)
    }
  } catch (error) {
    console.error('[ONLYOFFICE] 处理回调保存文件时出错:', error)
    // 按 ONLYOFFICE 协议，如果 error != 0，文档服务器会重试或认为保存失败
    // 这里仍返回 error:0，表示回调已处理完毕，避免无限重试
  }

  // 按 ONLYOFFICE 协议，返回 { error: 0 } 即表示保存成功
  return res.json({ error: 0 })
})

// 创建 HTTP 服务器
const server = createServer(app)

// 创建 WebSocket 服务器
const wss = new WebSocketServer({ server })

// 存储连接的客户端（插件和 Vue 应用）
const clients = {
  plugin: null, // 插件连接
  vue: null,    // Vue 应用连接
}

wss.on('connection', (ws, req) => {
  const url = new URL(req.url, `http://${req.headers.host}`)
  const clientType = url.searchParams.get('type') || 'unknown'

  console.log(`[WebSocket] 客户端连接: ${clientType}`)

  // 根据类型存储连接
  if (clientType === 'plugin') {
    clients.plugin = ws
  } else if (clientType === 'vue') {
    clients.vue = ws
  }

  // 处理消息
  ws.on('message', (message) => {
    try {
      const data = JSON.parse(message.toString())
      console.log(`[WebSocket] 收到 ${clientType} 消息:`, data)

      if (clientType === 'vue') {
        // Vue 应用发送的命令，转发给插件
        if (clients.plugin && clients.plugin.readyState === 1) {
          clients.plugin.send(JSON.stringify(data))
          console.log('[WebSocket] 命令已转发给插件')
        } else {
          // 插件未连接，返回错误
          ws.send(JSON.stringify({
            id: data.id,
            type: 'error',
            error: '插件未连接',
          }))
        }
      } else if (clientType === 'plugin') {
        // 插件返回的结果，转发给 Vue 应用
        if (clients.vue && clients.vue.readyState === 1) {
          clients.vue.send(JSON.stringify(data))
          console.log('[WebSocket] 结果已转发给 Vue 应用')
        }
      }
    } catch (error) {
      console.error('[WebSocket] 消息解析错误:', error)
    }
  })

  // 处理断开连接
  ws.on('close', () => {
    console.log(`[WebSocket] 客户端断开: ${clientType}`)
    if (clientType === 'plugin') {
      clients.plugin = null
    } else if (clientType === 'vue') {
      clients.vue = null
    }
  })

  // 发送连接成功消息
  ws.send(JSON.stringify({
    type: 'connected',
    clientType,
  }))
})

server.listen(PORT, () => {
  console.log(`[ONLYOFFICE] Callback demo server listening on http://localhost:${PORT}`)
  console.log(`[WebSocket] WebSocket server listening on ws://localhost:${PORT}`)
  console.log('[ONLYOFFICE] Static files served from /files/, e.g. /files/demo.docx')
  console.log('[ONLYOFFICE] All callbacks return {"error":0} and do NOT persist files.')
})

