// 测试脚本：监控文件变化，确认是否真的有自动保存
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const FILES_DIR = path.resolve(__dirname, 'files')
const ORIGINAL_FILE = path.join(FILES_DIR, 'demo.docx')
const LATEST_FILE = path.join(FILES_DIR, 'demo_latest.docx')

console.log('开始监控文件变化...')
console.log('原始文件:', ORIGINAL_FILE)
console.log('最新文件:', LATEST_FILE)
console.log('---')

// 记录初始状态
let originalStats = fs.existsSync(ORIGINAL_FILE) ? fs.statSync(ORIGINAL_FILE) : null
let latestStats = fs.existsSync(LATEST_FILE) ? fs.statSync(LATEST_FILE) : null

if (originalStats) {
    console.log('[初始] demo.docx 修改时间:', originalStats.mtime.toISOString())
}
if (latestStats) {
    console.log('[初始] demo_latest.docx 修改时间:', latestStats.mtime.toISOString())
}
console.log('---')

// 每秒检查一次
setInterval(() => {
    const newOriginalStats = fs.existsSync(ORIGINAL_FILE) ? fs.statSync(ORIGINAL_FILE) : null
    const newLatestStats = fs.existsSync(LATEST_FILE) ? fs.statSync(LATEST_FILE) : null

    // 检查原始文件是否被修改
    if (originalStats && newOriginalStats) {
        if (originalStats.mtime.getTime() !== newOriginalStats.mtime.getTime()) {
            console.log('⚠️ [警告] demo.docx 被修改了！')
            console.log('   旧时间:', originalStats.mtime.toISOString())
            console.log('   新时间:', newOriginalStats.mtime.toISOString())
            originalStats = newOriginalStats
        }
    }

    // 检查最新文件是否被修改
    if (latestStats && newLatestStats) {
        if (latestStats.mtime.getTime() !== newLatestStats.mtime.getTime()) {
            console.log('✓ demo_latest.docx 被更新了')
            console.log('   旧时间:', latestStats.mtime.toISOString())
            console.log('   新时间:', newLatestStats.mtime.toISOString())
            latestStats = newLatestStats
        }
    }
}, 1000)

console.log('监控中... (按 Ctrl+C 退出)')
