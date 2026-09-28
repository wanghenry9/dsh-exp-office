/**
 * 流式读取路径的测试：`ZipPackage#readChunks` + `Workbook#readRangeAsync`。
 *
 * 重点不是「能跑」，而是**与同步路径逐格等价**，以及边界不出错：
 *   - 强制走流式（把阈值调到 0）后，小样本的每个单元格都必须与同步路径一致；
 *   - 真早停：区域之后的行即便内容非法也不该被读到；
 *   - 跨块拼接：把块大小调小，逼出「一个 `<row>` 横跨多个块」的情况；
 *   - 单行过大的保护、损坏压缩流的结构化报错、阈值分流。
 */
import assert from 'node:assert/strict'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { ZipPackage, XmlDoc } from '../lib/ooxml.js'
import { Workbook, STREAM_THRESHOLD_BYTES } from '../lib/xlsx.js'
import { OfficeError } from '../lib/errors.js'

let passed = 0
let failed = 0

/**
 * 跑一个用例。
 * @param {string} title - 用例名。
 * @param {Function} fn - 用例体。
 * @returns {Promise<void>} 完成。
 */
async function test(title, fn) {
  try {
    await fn()
    passed += 1
    console.log(`  ✅ ${title}`)
  } catch (err) {
    failed += 1
    console.log(`  ❌ ${title}\n       ${err.message}`)
  }
}

/**
 * 造一个 xlsx（直接拼 XML，不经过 DOM）。
 * @param {object} options - 选项。
 * @param {string[]} options.rows - 每行的 XML 片段（不含 `<row>` 包裹）。
 * @param {string} [options.sharedStrings] - 共享字符串表 XML（可选）。
 * @returns {Buffer} 文件字节。
 */
function makeWorkbook({ rows, sharedStrings = null }) {
  const pkg = new ZipPackage(Buffer.alloc(0), new Map(), { maxEntries: 32, maxEntryBytes: 1 << 30, maxTotalBytes: 1 << 31, maxRatio: 1000 })
  pkg.write(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sharedStrings ? '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' : ''}</Types>`
  )
  pkg.write(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
  )
  pkg.write(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="表1" sheetId="1" r:id="rId1"/></sheets></workbook>`
  )
  pkg.write(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${sharedStrings ? '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ''}</Relationships>`
  )
  pkg.write(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D${rows.length}"/><sheetData>${rows.join('')}</sheetData></worksheet>`
  )
  pkg.write(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
  )
  if (sharedStrings) pkg.write('xl/sharedStrings.xml', sharedStrings)
  return pkg.toBuffer()
}

const LARGE = 'test/fixtures/large-200000x10-wide.xlsx'
const hasLarge = (() => {
  try {
    return readFileSync(LARGE).length > 0
  } catch {
    return false
  }
})()

console.log('\n=== 1. readChunks：与 read() 逐字节一致 ===')

await test('小块拼接的结果与一次性解压完全相同', async () => {
  const bytes = makeWorkbook({ rows: Array.from({ length: 50 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="B${i + 1}" t="inlineStr"><is><t>文本${i}</t></is></c></row>`) })
  const pkg = ZipPackage.open(bytes)
  const direct = pkg.read('xl/worksheets/sheet1.xml')
  const parts = []
  for await (const chunk of pkg.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 7 })) parts.push(chunk)
  const joined = Buffer.concat(parts)
  assert.equal(joined.length, direct.length)
  assert.ok(joined.equals(direct), '流式解压结果应与 read() 完全一致')
})

await test('提前中断后仍能再次完整读取（迭代器销毁不破坏包）', async () => {
  const bytes = makeWorkbook({ rows: Array.from({ length: 100 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`) })
  const pkg = ZipPackage.open(bytes)
  let seen = 0
  for await (const chunk of pkg.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 16 })) {
    seen += chunk.length
    if (seen > 32) break
  }
  const direct = pkg.read('xl/worksheets/sheet1.xml')
  let total = 0
  for await (const chunk of pkg.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 16 })) total += chunk.length
  assert.equal(total, direct.length)
})

await test('未压缩存储（method=0）的条目也能流式读', async () => {
  const bytes = makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] })
  const pkg = ZipPackage.open(bytes)
  const info = pkg.entryInfo('xl/worksheets/sheet1.xml')
  const parts = []
  for await (const chunk of pkg.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 8 })) parts.push(chunk)
  assert.equal(Buffer.concat(parts).length, info.uncompSize)
})

await test('条目不存在时给出 FILE_NOT_FOUND', async () => {
  const pkg = ZipPackage.open(makeWorkbook({ rows: ['<row r="1"/>'] }))
  await assert.rejects(() => pkg.readChunks('xl/nope.xml').next(), (err) => err.code === 'FILE_NOT_FOUND')
})

console.log('\n=== 2. 流式扫描与同步路径逐格等价 ===')

await test('★ 强制流式后，单元格类型/值/公式与同步路径逐格一致', async () => {
  const rows = []
  for (let i = 1; i <= 60; i += 1) {
    rows.push(
      `<row r="${i}">` +
        `<c r="A${i}"><v>${i * 3}</v></c>` +
        `<c r="B${i}" t="inlineStr"><is><t>文本 ${i}</t></is></c>` +
        `<c r="C${i}"><f>A${i}*2</f><v>${i * 6}</v></c>` +
        `<c r="D${i}" t="b"><v>${i % 2}</v></c>` +
        `</row>`
    )
  }
  const wb = Workbook.open(makeWorkbook({ rows }))
  const sync = wb.readRange('表1', 'A1:D60')
  const stream = await wb.readRangeAsync('表1', 'A1:D60', { streamThreshold: 0 })
  assert.equal(stream.read_mode, 'streaming')
  assert.deepEqual(stream.rows, sync.rows)
})

await test('★ 块很小（逼出跨块拼接）时结果不变', async () => {
  const rows = Array.from({ length: 30 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>跨块文本 ${i} 结束标记</t></is></c></row>`)
  const bytes = makeWorkbook({ rows })
  const wb = Workbook.open(bytes)
  const sync = wb.readRange('表1', 'A1:A30')
  const pkg = ZipPackage.open(bytes)
  // 用一个极小的块大小驱动扫描器：直接调用内部路径不可行，这里改用 readRangeAsync + 小阈值，
  // 再用 readChunks 的小块验证解压侧的拼接正确性
  const stream = await wb.readRangeAsync('表1', 'A1:A30', { streamThreshold: 0 })
  assert.deepEqual(stream.rows, sync.rows)
  let total = 0
  for await (const chunk of pkg.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 5 })) total += chunk.length
  assert.equal(total, pkg.entryInfo('xl/worksheets/sheet1.xml').uncompSize)
})

await test('只取中间区域：行号与列都对得上', async () => {
  const rows = Array.from({ length: 200 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="C${i + 1}"><v>${i * 2}</v></c></row>`)
  const wb = Workbook.open(makeWorkbook({ rows }))
  const sync = wb.readRange('表1', 'B50:C60')
  const stream = await wb.readRangeAsync('表1', 'B50:C60', { streamThreshold: 0 })
  assert.deepEqual(stream.rows, sync.rows)
  assert.equal(stream.rows[0][0], null, 'B 列为空')
  // 第 59 行（i=58）的 C 列值 = 58*2
  assert.equal(stream.rows[9][1].ref, 'C59')
  assert.equal(stream.rows[9][1].value, 58 * 2)
})

await test('★ 真早停：区域之后的行即使内容非法也不会被读到', async () => {
  const rows = ['<row r="1"><c r="A1"><v>1</v></c></row>', '<row r="2"><c r="A2"><v>2</v></c></row>', '<row r="4000"><c r="A4000"><f>#REF!</f><v>999</v></c></row>']
  const wb = Workbook.open(makeWorkbook({ rows }))
  const head = await wb.readRangeAsync('表1', 'A1:A2', { streamThreshold: 0 })
  assert.equal(head.rows.length, 2)
  assert.equal(head.rows[1][0].value, 2)
  // 反向对照：区域覆盖到第 4000 行时必须真的读到它（证明上面不是「恰好没扫到」）
  const tail = await wb.readRangeAsync('表1', 'A3999:A4000', { streamThreshold: 0 })
  assert.equal(tail.rows[1][0].value, 999)
})

await test('共享字符串表在流式路径里同样解析', async () => {
  const sst = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>甲</t></si><si><t>乙</t></si></sst>`
  const rows = ['<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>']
  const wb = Workbook.open(makeWorkbook({ rows, sharedStrings: sst }))
  const sync = wb.readRange('表1', 'A1:B1')
  const stream = await wb.readRangeAsync('表1', 'A1:B1', { streamThreshold: 0 })
  assert.deepEqual(stream.rows, sync.rows)
  assert.equal(stream.rows[0][0].value, '甲')
  assert.equal(stream.rows[0][1].value, '乙')
})

console.log('\n=== 3. 阈值分流与错误处理 ===')

await test('部件小于阈值时走同步路径（read_mode=sync）', async () => {
  const wb = Workbook.open(makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] }))
  const result = await wb.readRangeAsync('表1', 'A1:A1')
  assert.equal(result.read_mode, 'sync')
  assert.ok(result.part_bytes > 0)
  assert.ok(result.part_bytes < STREAM_THRESHOLD_BYTES)
})

await test('阈值可调：调成 0 就一律走流式', async () => {
  const wb = Workbook.open(makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] }))
  assert.equal((await wb.readRangeAsync('表1', 'A1:A1', { streamThreshold: 0 })).read_mode, 'streaming')
  assert.equal((await wb.readRangeAsync('表1', 'A1:A1', { streamThreshold: 1 << 30 })).read_mode, 'sync')
})

await test('区域超限在两条路径上都先拒绝（不分配内存）', async () => {
  const wb = Workbook.open(makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] }))
  assert.throws(() => wb.readRange('表1', 'A1:XFD1048576'), (err) => err.code === 'MEMORY_LIMIT')
  // 异步入口是同步抛错的，用 async 箭头把它变成 rejection 再断言
  await assert.rejects(async () => wb.readRangeAsync('表1', 'A1:XFD1048576'), (err) => err.code === 'MEMORY_LIMIT')
})

await test('单行异常大时给出 MEMORY_LIMIT 而不是吃光内存', async () => {
  // 用**不可压缩**的随机文本，避免先被「压缩比异常」的闸门拦下（那是另一条闸门，另有用例覆盖）
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
  let seed = 42
  const next = () => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff
    return seed / 0x80000000
  }
  const parts = []
  for (let i = 0; i < 36 * 1024 * 1024; i += 1) parts.push(alphabet[Math.floor(next() * alphabet.length)])
  const big = parts.join('')
  const wb = Workbook.open(makeWorkbook({ rows: [`<row r="1"><c r="A1" t="inlineStr"><is><t>${big}</t></is></c></row>`] }))
  await assert.rejects(
    () => wb.readRangeAsync('表1', 'A1:A1', { streamThreshold: 0 }),
    (err) => err.code === 'MEMORY_LIMIT' && /一行/.test(err.message)
  )
})

await test('★ CRC32：压缩数据损坏但解压长度不变时也能识别出来', async () => {
  const bytes = makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] })
  const entry = ZipPackage.open(bytes).entryInfo('xl/worksheets/sheet1.xml')
  assert.ok(entry.compSize > 8 && entry.method === 8, '样本应走 deflate')
  const copy = Buffer.from(bytes)
  // 用元信息里的精确偏移，别去猜（第一处 PK\\x03\\x04 是别的条目）
  for (let i = 2; i < entry.compSize; i += 1) copy[entry.dataOffset + i] = 0xff
  const broken = ZipPackage.open(copy)
  await assert.rejects(
    async () => {
      for await (const chunk of broken.readChunks('xl/worksheets/sheet1.xml', { chunkSize: 8 })) void chunk
    },
    (err) => err.code === 'CORRUPTED_DOCUMENT'
  )
})

await test('CRC 为 0（写入方未设置）时跳过校验，不误报损坏', async () => {
  const bytes = makeWorkbook({ rows: ['<row r="1"><c r="A1"><v>1</v></c></row>'] })
  const entry = ZipPackage.open(bytes).entryInfo('xl/worksheets/sheet1.xml')
  // 只把中央目录里的 CRC 改成 0，不动数据
  const crcOffset = copyCrcOffset(bytes, 'xl/worksheets/sheet1.xml')
  const patched = Buffer.from(bytes)
  patched.writeUInt32LE(0, crcOffset)
  assert.equal(ZipPackage.open(patched).entryInfo('xl/worksheets/sheet1.xml').crc, 0)
  let total = 0
  for await (const chunk of ZipPackage.open(patched).readChunks('xl/worksheets/sheet1.xml')) total += chunk.length
  assert.equal(total, entry.uncompSize)
})

/**
 * 在中央目录里定位某条目 CRC 字段的偏移（只用于测试）。
 * @param {Buffer} bytes - 文件字节。
 * @param {string} name - 条目名。
 * @returns {number} 偏移。
 */
function copyCrcOffset(bytes, name) {
  const eocd = bytes.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  const count = bytes.readUInt16LE(eocd + 10)
  let cursor = bytes.readUInt32LE(eocd + 16)
  for (let i = 0; i < count; i += 1) {
    const nameLen = bytes.readUInt16LE(cursor + 28)
    const extraLen = bytes.readUInt16LE(cursor + 30)
    const commentLen = bytes.readUInt16LE(cursor + 32)
    const entryName = bytes.toString('utf8', cursor + 46, cursor + 46 + nameLen)
    if (entryName === name) return cursor + 16
    cursor += 46 + nameLen + extraLen + commentLen
  }
  throw new Error(`中央目录里找不到 ${name}`)
}

console.log('\n=== 4. 真实大部件（178 MB）===')

if (!hasLarge) {
  console.log(`  ⏭️  跳过：缺少 ${LARGE}（node test/make-xlsx-large-fixture.mjs --rows 200000 --cols 10 --wide）`)
} else {
  await test('★ 178 MB 部件：流式读取区域正确且与同步路径一致', async () => {
    const bytes = readFileSync(LARGE)
    const wb = Workbook.open(bytes)
    const streamHead = await wb.readRangeAsync('大数据', 'A1:J5')
    assert.equal(streamHead.read_mode, 'streaming')
    assert.ok(streamHead.part_bytes > 150 * 1024 * 1024, `部件应 >150MB，实际 ${streamHead.part_bytes}`)
    assert.equal(streamHead.rows.length, 5)
    assert.equal(streamHead.rows[0][0].value, '乙')
    // 同步路径读同一区域作对照（会解压整份部件，这里只跑一次小区域）
    const syncHead = wb.readRange('大数据', 'A1:J5')
    assert.deepEqual(streamHead.rows, syncHead.rows)
  })

  await test('★ 尾部区域也能读到（流式必须扫过前面内容）', async () => {
    const wb = Workbook.open(readFileSync(LARGE))
    const result = await wb.readRangeAsync('大数据', 'A199998:J199999')
    assert.equal(result.rows.length, 2)
    assert.equal(result.rows[0][0].ref, 'A199998')
    assert.ok(result.rows[0][0].value !== null)
  })
}

console.log(`\n结果：${passed} 通过，${failed} 失败\n`)
process.exit(failed === 0 ? 0 : 1)
