# Harness 版本兼容矩阵

> 目的：把「声明支持 `>=0.1.0-rc.6`」从**一句话**变成**可复现的实测数据**。
> 结论先行：**5 个 `dsh-tools` 版本 × 3 个 DSH 发行版全部通过**，其中包含宿主版本与当时最新的 `0.1.7-rc.2`。

## 为什么兼容面只有两样东西

本插件的 `lib/` **不 import 任何 `@deepseek-ai/*` 包**（零运行时依赖）：

```bash
grep -rn "from '@deepseek" lib/    # 无输出
```

`defineTool` 是宿主通过 `apply(ctx, { defineTool })` **注入**进来的，Cordis 的 `Context` 由宿主提供。
所以需要验证的兼容面就是：

1. 那个版本的 **`@deepseek-ai/dsh-tools`**：能不能定义我们的 90 个工具（值 schema DSL）、调用、渲染；
2. 那个版本的 **`@deepseek-ai/cordis`**：`apply` 的 effect 契约与 `ctx.plugin()` 装载协议；
3. 那个版本的 **DSH CLI / 加载器**：接不接受我们的 `dsh.plugin.json` 与 `cordis.patch.yml`。

## 第一层：工具定义与调用（`npm run verify:matrix`）

对每个版本：临时目录里装那个版本的 `@deepseek-ai/dsh-tools`（它自带匹配的 `cordis`）→
用它的 `defineTool` 跑 `apply()` → 检查定义形状 → 用真实 Cordis `Context` 装载 → 跑三个真实调用
（生成 PDF / 列目录 / 坏输入必须结构化报错）→ 再用那个版本的渲染契约渲染一次。

| dsh-tools | cordis | 定义形状 | effect | 装载 | 调用+渲染 | 结论 |
|---|---|---|---|---|---|---|
| `0.1.2-rc.1` | 4.0.4 | ✓ 90 | ✓ | 90 | ✓ | ✅ 兼容 |
| `0.1.5-rc.1` | 4.0.2 | ✓ 90 | ✓ | 90 | ✓ | ✅ 兼容 |
| `0.1.5-rc.3` | 4.0.2 | ✓ 90 | ✓ | 90 | ✓ | ✅ 兼容 |
| `0.1.6-alpha.2`（宿主） | 4.0.4 | ✓ 90 | ✓ | 90 | ✓ | ✅ 兼容 |
| `0.1.7-rc.2`（当时最新） | 4.0.4 | ✓ 90 | ✓ | 90 | ✓ | ✅ 兼容 |

**「定义形状」这一列不是形式主义**：`dsh-tools@0.1.7-rc.2` 把 `output.render` 从可选变成了**必需**
（实现里直接读 `options.output.render`，`output` 缺失时抛 `Cannot read properties of undefined (reading 'render')`）。
我们的 `envelopeOutput()` 一直带着 `render`，因此两个版本都通过 —— 若当初少写这一项，插件在 0.1.7 上会直接挂掉，
而这类问题**只在跨版本测试里才会暴露**。

## 第二层：真实 CLI 装一遍（`npm run verify:matrix:cli`）

对每个 DSH 发行版：临时目录装**那个版本的 DSH 应用本身** → 用**它自己的 CLI**、
在**独立的临时 `DSH_HOME`** 里走一遍用户路径：

```bash
dsh plugin --profile matrix add github:wanghenry9/dsh-exp-office   # 装进 profile
dsh --profile matrix --dump-config                                 # 组合配置树
```

断言四件事：CLI 报的版本号与预期一致、插件真的进了 profile 依赖、被注册进 `dsh.profile.bundles`、
组合后的配置树里出现我们的补丁层（`# == dsh-exp-office` / `- id: dsh-exp-office`）。

| DSH 版本 | CLI 报的版本 | 装进 profile | 注册为 bundle | 组合树含插件层 | 结论 |
|---|---|---|---|---|---|
| `0.1.6-alpha.2`（宿主） | `0.1.6-alpha.2` | ✓ | ✓ | ✓ `# == dsh-exp-office` | ✅ 兼容 |
| `0.1.7-rc.2`（当时最新） | `0.1.7-rc.2` | ✓ | ✓ | ✓ `# == dsh-exp-office` | ✅ 兼容 |

```text
[tree] 0.1.7-rc.2: # == dsh-exp-office | - id: dsh-exp-office | name: dsh-exp-office
```

**更早的发行版不再默认测**：`dsh@0.1.1-rc.1` 的依赖树在本机装 20 分钟仍无产出（npm 进程退出、
目录还是空的），属于上游旧版本的安装问题，与我们无关。需要时显式传版本号并放宽超时：

```bash
node test/harness-cli-matrix.mjs 0.1.1-rc.1 --install-timeout 1200000
```

> 顺带修掉了一个脚本自身的坑：最初用管道收集子进程输出，npm 留下的孙进程**持有管道写端**，
> `execFileSync` 会一直等 EOF —— 表现就是「卡住不返回」。现在改成**把输出重定向到文件**，
> 并给每一步加超时（安装 5 分钟、CLI 调用 5 分钟，可命令行放宽）。

## 怎么复现

```bash
npm run verify:matrix              # 第一层，约 1 分钟（需要网络：装 dsh-tools 各版本）
npm run verify:matrix:cli          # 第二层，首次约 5–6 分钟（每个版本下载 453 MB 应用）
node test/harness-matrix.mjs 0.1.7-rc.2 0.1.6-alpha.2      # 指定版本（第一层）
node test/harness-cli-matrix.mjs 0.1.7-rc.2                # 指定版本（第二层）
node test/harness-cli-matrix.mjs --source file             # 用本地工作区代替 GitHub
```

两个脚本都不进 `npm test`：它们需要网络。`npm test` 必须保持离线可跑。

## 隔离与清理

- 第二层每个版本用**独立的临时 `DSH_HOME`**，跑完即删 —— **绝不碰用户正在用的 `~/.dsh`**
  （脚本会显式覆盖 `DSH_HOME`，并在开始前把它从继承环境里剔除）；
- 下载好的 DSH 应用缓存在系统临时目录的 `dsh-matrix-cache-<版本>`，重跑不必再下载；
- 临时目录名一律带 `dsh-matrix-` / `dsh-home-cli-` 前缀，避免误删宿主自己的 `dsh-*` 运行目录。

## 覆盖边界（诚实说明）

- **已覆盖**：工具层从 `0.1.2-rc.1`（npm 上最早可安装的 `dsh-tools`）到 `0.1.7-rc.2`，三种 cordis 组合（4.0.2 / 4.0.4）；CLI 层覆盖宿主 `0.1.6-alpha.2` 与最新 `0.1.7-rc.2`。
- **未覆盖**：声明里的 `0.1.0-rc.6`〜`0.1.1-rc.2` —— 这些版本在 npm 上没有可安装的
  `dsh-tools`（`@deepseek-ai/dsh` 也最早只到 `0.1.1-rc.1`），**无法取得用于验证**；
  见下表的实测下限。
- **未覆盖**：真实的会话内调用（需要模型与 API 额度）。这里验证的是装载、定义与本地调用契约，
  不包含「Agent 在对话里点得动这些工具」—— 那属于宿主 GUI 的验证范围。
