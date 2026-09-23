/**
 * 把工作区同步到 profile 的安装副本 —— **不调用 pnpm**。
 *
 * 为什么不用 pnpm：`pnpm install` 会按 package.json **收敛整棵 node_modules**，
 * 安装非本包的依赖时还会去替换被 DSH 进程占用的原生包（实测报
 * `failed to remove existing directory … 拒绝访问 (os error 5)` 并以退出码 1 结束）。
 * 这会影响 profile 里**其它插件**的依赖树 —— 属于「改别人的东西」。
 *
 * 这个脚本只做一件事：把本包的 lib/ 与三个清单文件复制到
 * `<profile>/node_modules/dsh-exp-office/`，其它目录一律不碰。
 *
 * 用法：node test/sync-to-profile.mjs [profileDir]
 */
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync } from 'node:fs'
import os from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const workspace = join(here, '..')
const profileDir = process.argv[2] ?? `${os.homedir()}/.dsh/profiles/web`
const target = join(profileDir, 'node_modules', 'dsh-exp-office')

if (!existsSync(profileDir)) {
  console.log(`⏭️  跳过：profile 目录不存在 ${profileDir}`)
  process.exit(0)
}
if (!existsSync(target)) {
  console.log(`❌ 安装副本不存在：${target}\n   首次安装仍需在 profile 目录执行一次 pnpm install（建议在 DSH 停止时进行）。`)
  process.exit(1)
}

let copied = 0
mkdirSync(join(target, 'lib'), { recursive: true })
for (const name of readdirSync(join(workspace, 'lib'))) {
  if (!name.endsWith('.js')) continue
  const from = join(workspace, 'lib', name)
  const to = join(target, 'lib', name)
  if (existsSync(to) && readFileSync(from).compare(readFileSync(to)) === 0) continue
  copyFileSync(from, to)
  copied += 1
}
for (const name of ['dsh.plugin.json', 'cordis.patch.yml', 'package.json']) {
  const from = join(workspace, name)
  const to = join(target, name)
  if (!existsSync(from)) continue
  if (existsSync(to) && readFileSync(from).compare(readFileSync(to)) === 0) continue
  copyFileSync(from, to)
  copied += 1
}
console.log(`✅ 已同步 ${copied} 个文件到 ${target}（未触碰 profile 的其它依赖）`)
