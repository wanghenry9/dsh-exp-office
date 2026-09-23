/**
 * 推前隐私扫描（发布闸门之一）。
 *
 * 与「敏感信息扫描」（密钥类）互补：这一关专门盯**个人信息与环境路径**。
 * 设计上刻意把「要藏的东西」放在**不入库**的本地文件里：
 *
 *   - `.privacy-patterns.local`（已 gitignore）：一行一个模式，可以写 `#` 注释。
 *     推前扫描会把每个模式当**字面量**在已跟踪文件里查找。
 *     这个文件本身绝不入库 —— 否则等于把要藏的东西写进仓库。
 *   - 内置规则（无需配置，且不含任何个人信息）：Windows 用户目录绝对路径、
 *     GitHub noreply 邮箱、本机用户名出现在路径里。
 *
 * 扫描范围：有 git 仓库时只扫**已跟踪/已暂存**文件（`git ls-files`），
 * 否则退回扫描工作区（排除 node_modules / .git / .dsh-release 与二进制文件）。
 *
 * 退出码：0 = 干净；3 = 有 BLOCKER；2 = 用法错误。
 *
 * 用法：
 *   node test/release-privacy-check.mjs            # 全部已跟踪文件
 *   node test/release-privacy-check.mjs --staged   # 只看已暂存文件（推前最后一道）
 */
import fs from 'node:fs'
import path from 'node:path'
import { execFileSync } from 'node:child_process'

const ROOT = path.resolve(import.meta.dirname, '..')
const PATTERN_FILE = path.join(ROOT, '.privacy-patterns.local')
const stagedOnly = process.argv.includes('--staged')

/** 内置规则：与具体个人无关，所以可以安全地写在代码里。 */
const BUILT_IN = [
  { name: 'Windows 用户目录绝对路径', regex: /[A-Za-z]:[\\/]Users[\\/][^\\/\s"']+/g, hint: '改成 %USERPROFILE% 或 os.homedir()' },
  { name: 'GitHub noreply 邮箱', regex: /[\w.+-]+@users\.noreply\.github\.com/g, hint: '不要提交个人邮箱' },
  { name: '疑似本机用户名出现在 .dsh 路径', regex: /Users[\\/][^\\/\s"']+[\\/]\.dsh/g, hint: '改成 %USERPROFILE%\\.dsh' }
]

const SKIP_DIRS = new Set(['node_modules', '.git', '.dsh-release', 'dist', 'build'])
const TEXT_EXT = /\.(js|mjs|cjs|ts|json|md|txt|yml|yaml|ps1|sh|xml|rels|html|css|gitignore|gitattributes)$/i

/**
 * 读本地模式表。
 * @returns {string[]} 模式列表（已去掉注释与空行）。
 */
function readPatterns() {
  if (!fs.existsSync(PATTERN_FILE)) return []
  return fs
    .readFileSync(PATTERN_FILE, 'utf8')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== '' && !line.startsWith('#'))
}

/**
 * 列出要扫描的文件。
 * @returns {string[]} 相对 ROOT 的文件路径。
 */
function listFiles() {
  try {
    const args = stagedOnly ? ['diff', '--cached', '--name-only', '--diff-filter=ACMR'] : ['ls-files']
    const out = execFileSync('git', args, { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] })
    return out.split(/\r?\n/).filter((line) => line !== '')
  } catch {
    // 还没有 git 仓库（或没装 git）：退回遍历工作区
    const found = []
    const walk = (dir) => {
      for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        if (entry.name.startsWith('.') && entry.name !== '.gitignore' && entry.name !== '.gitattributes') continue
        if (SKIP_DIRS.has(entry.name)) continue
        const full = path.join(dir, entry.name)
        if (entry.isDirectory()) walk(full)
        else found.push(path.relative(ROOT, full))
      }
    }
    walk(ROOT)
    return found
  }
}

const patterns = readPatterns()
const files = listFiles()
const rules = [...BUILT_IN.map((r) => ({ ...r, patterns: null })), ...(patterns.length ? [{ name: '本地隐私模式', patterns, hint: '来自 .privacy-patterns.local' }] : [])]

const blockers = []
for (const rel of files) {
  if (!TEXT_EXT.test(rel)) continue
  const full = path.join(ROOT, rel)
  if (!fs.existsSync(full)) continue
  const text = fs.readFileSync(full, 'utf8')
  for (const rule of rules) {
    if (rule.patterns) {
      for (const pattern of rule.patterns) {
        let index = text.indexOf(pattern)
        while (index >= 0) {
          const line = text.slice(0, index).split(/\r?\n/).length
          blockers.push({ file: rel, line, rule: rule.name, hint: rule.hint })
          index = text.indexOf(pattern, index + pattern.length)
        }
      }
      continue
    }
    for (const match of text.matchAll(rule.regex)) {
      const line = text.slice(0, match.index).split(/\r?\n/).length
      // 占位写法（%USERPROFILE%、<用户名>）是允许的，不算命中
      if (/%USERPROFILE%|<用户名>|<user>|<name>/.test(match[0])) continue
      blockers.push({ file: rel, line, rule: rule.name, hint: rule.hint, sample: match[0] })
    }
  }
}

console.log('=== 推前隐私扫描 ===')
console.log(`模式来源: ${patterns.length ? `.privacy-patterns.local（${patterns.length} 条，内容不打印）` : '无本地模式表（仅内置规则）'}`)
console.log(`扫描文件: ${files.length} 个${stagedOnly ? '（仅已暂存）' : ''}`)
if (!patterns.length) {
  console.log('提示: 未找到 .privacy-patterns.local —— 若这台机器上有需要回避的个人标识，请把模式写进该文件（它不会入库）。')
}
if (blockers.length === 0) {
  console.log('结论: 干净，未发现个人信息或本机路径。')
  process.exit(0)
}
console.log(`结论: BLOCKER=${blockers.length}，必须先处理：`)
for (const item of blockers) {
  console.log(`  [!] ${item.file}:${item.line}  [${item.rule}]${item.sample ? `  → ${item.sample}` : ''}  （${item.hint}）`)
}
process.exit(3)
