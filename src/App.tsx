import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'
import catStickerSheet from './assets/cat-sticker-sheet.png'
import { supabase } from './supabase'

type Member = {
  id: string
  name: string
  power: number
  previousPower?: number
  weeklyPower: number
  score: number | null
  remark: string
  order: number
}

type PackageType = 'fire' | 'mid1' | 'mid2'
type Schedule = Record<string, string | null>
type AccessoryName = '手镯' | '戒指' | '耳环' | '腰带' | '项链' | '徽章' | '剩余饰品'
type QueueEntry = { id: string; name: string; addedAt: string }
type DistributionCounts = Record<AccessoryName, number>
type AccessoryDefinition = { name: AccessoryName; icon: string; color: string }

type Session = {
  id: string
  label: string
  block: 'pink' | 'orange' | 'yellow' | 'green' | 'cyan' | 'blue'
  startColumn: number
  startRow: number
}

const PACKAGE_LABELS: Record<PackageType, string> = { fire: '火', mid1: '中一', mid2: '中二' }
const TYPE_ORDER: PackageType[] = ['fire', 'mid1', 'mid2']
const ACCESSORIES: AccessoryDefinition[] = [
  { name: '手镯', icon: '🪬', color: 'pink' },
  { name: '戒指', icon: '💍', color: 'yellow' },
  { name: '耳环', icon: '✨', color: 'cyan' },
  { name: '腰带', icon: '🎀', color: 'green' },
  { name: '项链', icon: '📿', color: 'orange' },
  { name: '徽章', icon: '🏵️', color: 'blue' },
]
const QUEUE_ACCESSORIES: AccessoryDefinition[] = [...ACCESSORIES, { name: '剩余饰品', icon: '🎁', color: 'remainder' }]
const TIERS = [
  { min: 1, max: 5, fire: 2, middle: 3, coins: 62, color: 'pink' },
  { min: 6, max: 10, fire: 2, middle: 2, coins: 60, color: 'orange' },
  { min: 11, max: 15, fire: 2, middle: 1, coins: 58, color: 'yellow' },
  { min: 16, max: 20, fire: 1, middle: 3, coins: 55, color: 'green' },
  { min: 21, max: 25, fire: 1, middle: 2, coins: 53, color: 'cyan' },
  { min: 26, max: 30, fire: 0, middle: 5, coins: 52, color: 'blue' },
] as const

const SESSIONS: Session[] = [
  { id: 'sat-pm', label: '周六下', block: 'pink', startColumn: 1, startRow: 2 },
  { id: 'sun', label: '周日', block: 'pink', startColumn: 4, startRow: 2 },
  { id: 'mon', label: '周一', block: 'orange', startColumn: 7, startRow: 2 },
  { id: 'tue', label: '周二', block: 'green', startColumn: 10, startRow: 2 },
  { id: 'wed', label: '周三', block: 'yellow', startColumn: 1, startRow: 9 },
  { id: 'thu', label: '周四', block: 'yellow', startColumn: 4, startRow: 9 },
  { id: 'fri', label: '周五', block: 'cyan', startColumn: 7, startRow: 9 },
  { id: 'sat-am', label: '星期六上', block: 'green', startColumn: 10, startRow: 9 },
]

function emptyQueues(): Record<AccessoryName, QueueEntry[]> {
  return { 手镯: [], 戒指: [], 耳环: [], 腰带: [], 项链: [], 徽章: [], 剩余饰品: [] }
}

function emptyDistributionCounts(): DistributionCounts {
  return { 手镯: 0, 戒指: 0, 耳环: 0, 腰带: 0, 项链: 0, 徽章: 0, 剩余饰品: 0 }
}

function sweepAccessoryQueues(source: Record<AccessoryName, QueueEntry[]>, counts: DistributionCounts) {
  return QUEUE_ACCESSORIES.reduce((next, accessory) => {
    next[accessory.name] = source[accessory.name].slice(counts[accessory.name])
    return next
  }, emptyQueues())
}

function sundayDateKey() {
  const now = new Date()
  return `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()}`
}

const SAMPLE_NAMES = ['太初星', '关注塔菲喵', '花云青', '无双', '御茨星', '念君夏', '夏弥', '谦灵星', '小星星', '沈七涵', '晓行星', '心芯星', '猫猫星', '绝地', '鸿鹄', '云岫', '亿丈龙我', '椛七', '超级萝卜大王', '拳王', '浅帐星', '别急稳一手', '小苏在这里', '时愿星', '伦敦街尾吻别', '33', '季时雨花知否', '我一直都在', '弦', '白慕']
const SAMPLE_POWER = [3852, 3751, 3736, 3258, 3220, 3077, 2867, 2881, 2802, 2746, 2732, 2684, 2633, 2567, 2547, 2542, 2469, 2460, 2408, 2405, 2323, 2314, 2306, 2294, 2218, 2206, 2193, 2189, 2110, 2096]
const SAMPLE_SCORE_BY_NAME: Record<string, number> = {
  '太初星': 36, '关注塔菲喵': 36, '花云青': 36, '无双': 36, '御茨星': 35, '念君夏': 36, '夏弥': 36, '谦灵星': 36, '小星星': 35, '沈七涵': 16, '晓行星': 36, '心芯星': 36, '猫猫星': 36, '绝地': 36, '鸿鹄': 26, '云岫': 26, '亿丈龙我': 36, '椛七': 36, '超级萝卜大王': 5, '拳王': 36, '浅帐星': 36, '别急稳一手': 36, '小苏在这里': 15, '时愿星': 17, '伦敦街尾吻别': 36, '33': 36, '季时雨花知否': 36, '我一直都在': 15, '弦': 26, '白慕': 36,
}

function makeSampleMembers(): Member[] {
  return SAMPLE_NAMES.map((name, index) => ({ id: `m-${index + 1}`, name, power: SAMPLE_POWER[index], previousPower: SAMPLE_POWER[index], weeklyPower: 0, score: SAMPLE_SCORE_BY_NAME[name] ?? null, remark: '', order: index }))
}

function normalizeMember(member: Member): Member {
  const weeklyPower = Number(member.weeklyPower) || 0
  return { ...member, previousPower: member.previousPower ?? Math.max((Number(member.power) || 0) - weeklyPower, 0), weeklyPower }
}

function cloneMembers(source: Member[]) {
  return source.map((member) => ({ ...member }))
}

function tierForRank(rank: number) { return TIERS.find((tier) => rank >= tier.min && rank <= tier.max) ?? TIERS[TIERS.length - 1] }
function rankMembers(members: Member[]) { return [...members].sort((a, b) => (b.score ?? -1) - (a.score ?? -1) || b.power - a.power || a.order - b.order) }
function powerRankMembers(members: Member[]) { return [...members].sort((a, b) => b.power - a.power || a.order - b.order) }
function cellKey(sessionId: string, type: PackageType, row: number) { return `${sessionId}:${type}:${row}` }
function setGroup(schedule: Schedule, sessionId: string, type: PackageType, ids: (string | null)[]) { ids.forEach((id, row) => { schedule[cellKey(sessionId, type, row)] = id }) }

function buildAutoSchedule(ranked: Member[]): Schedule {
  const schedule: Schedule = {}
  SESSIONS.forEach((session) => TYPE_ORDER.forEach((type) => setGroup(schedule, session.id, type, [null, null, null, null, null])))
  const group = (from: number) => ranked.slice(from, from + 5).map((member) => member?.id ?? null)
  setGroup(schedule, 'sat-pm', 'fire', group(0)); setGroup(schedule, 'sat-pm', 'mid1', group(0)); setGroup(schedule, 'sat-pm', 'mid2', group(0))
  setGroup(schedule, 'sun', 'fire', group(0)); setGroup(schedule, 'sun', 'mid1', group(0)); setGroup(schedule, 'sun', 'mid2', group(10))
  setGroup(schedule, 'mon', 'fire', group(5)); setGroup(schedule, 'mon', 'mid1', group(5)); setGroup(schedule, 'mon', 'mid2', group(5))
  setGroup(schedule, 'tue', 'fire', group(5)); setGroup(schedule, 'tue', 'mid1', group(15)); setGroup(schedule, 'tue', 'mid2', group(25))
  setGroup(schedule, 'wed', 'fire', group(10)); setGroup(schedule, 'wed', 'mid1', group(25)); setGroup(schedule, 'wed', 'mid2', group(25))
  setGroup(schedule, 'thu', 'fire', group(10)); setGroup(schedule, 'thu', 'mid1', group(25)); setGroup(schedule, 'thu', 'mid2', group(25))
  setGroup(schedule, 'fri', 'fire', group(20)); setGroup(schedule, 'fri', 'mid1', group(20)); setGroup(schedule, 'fri', 'mid2', group(20))
  setGroup(schedule, 'sat-am', 'fire', group(15)); setGroup(schedule, 'sat-am', 'mid1', group(15)); setGroup(schedule, 'sat-am', 'mid2', group(15))
  return schedule
}

function parseCell(value: unknown): string { return value == null ? '' : String(value).trim() }
function numeric(value: unknown): number | null { const n = Number(value); return Number.isFinite(n) ? n : null }

async function importWorkbook(file: File): Promise<Member[]> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await file.arrayBuffer())
  const byName = new Map<string, Partial<Member> & { order: number }>()
  let order = 0
  const absorb = (name: unknown, power: unknown, score: unknown, remark: unknown) => {
    const clean = parseCell(name)
    if (!clean || clean === '合计' || clean === '人员') return
    const existing = byName.get(clean) ?? { order: order++ }
    const p = numeric(power); const s = numeric(score)
    if (p != null) existing.power = p
    if (s != null) existing.score = s
    if (parseCell(remark)) existing.remark = parseCell(remark)
    byName.set(clean, existing)
  }
  const readRankSheet = (sheet: ExcelJS.Worksheet, mode: 'power' | 'score') => {
    let header = -1; let nameCol = -1; let valueCol = -1; let remarkCol = -1
    sheet.eachRow((row, rowNumber) => {
      const values = row.values as unknown[]
      values.forEach((value, col) => { if (parseCell(value) === '人员') { header = rowNumber; nameCol = col } if (parseCell(value) === (mode === 'power' ? '战力' : '分数')) valueCol = col; if (parseCell(value) === '备注') remarkCol = col })
    })
    if (header < 0 || nameCol < 0) return
    for (let r = header + 1; r <= sheet.rowCount; r += 1) { const row = sheet.getRow(r); absorb(row.getCell(nameCol).value, mode === 'power' ? row.getCell(valueCol).value : undefined, mode === 'score' ? row.getCell(valueCol).value : undefined, remarkCol > 0 ? row.getCell(remarkCol).value : undefined) }
  }
  const powerSheet = workbook.getWorksheet('战力排名')
  const scoreSheet = workbook.getWorksheet('分数排名')
  if (powerSheet) readRankSheet(powerSheet, 'power')
  if (scoreSheet) readRankSheet(scoreSheet, 'score')
  const first = workbook.worksheets[0]
  if (first) {
    for (let r = 1; r <= first.rowCount; r += 1) {
      const row = first.getRow(r)
      for (let c = 1; c <= row.cellCount; c += 1) {
        if (parseCell(row.getCell(c).value) === '人员' && parseCell(row.getCell(c + 1).value) === '分数') {
          for (let rr = r + 1; rr <= Math.min(r + 31, first.rowCount); rr += 1) absorb(first.getRow(rr).getCell(c).value, undefined, first.getRow(rr).getCell(c + 1).value, first.getRow(rr).getCell(c + 2).value)
        }
      }
    }
  }
  return [...byName.entries()].map(([name, data], index) => ({ id: `m-${Date.now()}-${index}`, name, power: data.power ?? 0, weeklyPower: 0, score: data.score ?? null, remark: data.remark ?? '', order: data.order ?? index })).sort((a, b) => a.order - b.order)
}

function colorHex(color: string) { return ({ pink: 'F7DDE3', orange: 'FFC000', yellow: 'FFF200', green: '92D050', cyan: '10B8E8', blue: '4472C4' } as Record<string, string>)[color] ?? 'FFFFFF' }

function buildExportWorkbook(members: Member[], schedule: Schedule, ranked: Member[]) {
  const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet('繁星本周要塞包分配')
  sheet.properties.defaultRowHeight = 20
  for (let col = 1; col <= 12; col += 1) sheet.getColumn(col).width = 14
  sheet.getColumn(13).width = 8; sheet.getColumn(14).width = 22; sheet.getColumn(15).width = 10; sheet.getColumn(16).width = 20
  sheet.getColumn(17).width = 12; sheet.getColumn(18).width = 20; sheet.getColumn(19).width = 14; sheet.getColumn(20).width = 14; sheet.getColumn(21).width = 12
  const border = { top: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, left: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, bottom: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } }, right: { style: 'thin' as const, color: { argb: 'FF3F3F3F' } } }
  const center = { vertical: 'middle' as const, horizontal: 'center' as const }
  sheet.mergeCells('A1:U1'); sheet.getCell('A1').value = '繁星本周要塞包分配'; sheet.getCell('A1').font = { name: '宋体', size: 16, bold: true }; sheet.getCell('A1').alignment = center; sheet.getRow(1).height = 34
  const rankHeader = ['本周', '人员', '分数', '备注']; ['M2', 'N2', 'O2', 'P2'].forEach((cell, i) => { sheet.getCell(cell).value = rankHeader[i]; sheet.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7E6E6' } }; sheet.getCell(cell).font = { name: '宋体', size: 12 }; sheet.getCell(cell).alignment = center; sheet.getCell(cell).border = border })
  const powerRankHeader = ['战力排名', '成员姓名', '上周战力', '本周战力', '提升']; ['Q2', 'R2', 'S2', 'T2', 'U2'].forEach((cell, i) => { sheet.getCell(cell).value = powerRankHeader[i]; sheet.getCell(cell).font = { name: '宋体', size: 12, bold: true }; sheet.getCell(cell).alignment = center; sheet.getCell(cell).border = border })
  const memberById = new Map(members.map((member) => [member.id, member]))
  SESSIONS.forEach((session) => {
    const start = session.startColumn; const end = start + 2; const headerRow = session.startRow; const dataStart = headerRow + 2
    sheet.mergeCells(headerRow, start, headerRow, end); sheet.getCell(headerRow, start).value = session.label; sheet.getCell(headerRow, start).alignment = center; sheet.getCell(headerRow, start).font = { name: '楷体', size: 12 }; sheet.getCell(headerRow, start).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }
    TYPE_ORDER.forEach((type, offset) => { const assigned = memberById.get(schedule[cellKey(session.id, type, 0)] ?? ''); const assignedRank = assigned ? ranked.findIndex((entry) => entry.id === assigned.id) + 1 : 0; const cell = sheet.getCell(headerRow + 1, start + offset); cell.value = PACKAGE_LABELS[type]; cell.alignment = center; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(assignedRank ? tierForRank(assignedRank).color : session.block)}` } } })
    for (let row = 0; row < 5; row += 1) {
      const values = TYPE_ORDER.map((type) => memberById.get(schedule[cellKey(session.id, type, row)] ?? '')?.name ?? '')
      TYPE_ORDER.forEach((type, offset) => { const cell = sheet.getCell(dataStart + row, start + offset); cell.value = values[offset]; cell.alignment = center; cell.border = border; const member = memberById.get(schedule[cellKey(session.id, type, row)] ?? ''); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${rank ? colorHex(tierForRank(rank).color) : 'FFFFFF'}` } } })
      if (values[0] && values[0] === values[1] && values[1] === values[2]) { sheet.mergeCells(dataStart + row, start, dataStart + row, end); sheet.getCell(dataStart + row, start).value = values[0] } else if (values[0] && values[0] === values[1]) { sheet.mergeCells(dataStart + row, start, dataStart + row, start + 1); sheet.getCell(dataStart + row, start).value = values[0] }
    }
  })
  sheet.mergeCells('A17:L17'); sheet.getCell('A17').value = '分配方案：每周1火2中，每人每周可得42个币，另外火包7币，中包2币，按每周考核分来分包'; sheet.getCell('A17').alignment = center; sheet.getCell('A17').font = { name: '宋体', size: 11 }
  TIERS.forEach((tier, index) => { const row = 18 + index; sheet.mergeCells(row, 1, row, 3); sheet.mergeCells(row, 4, row, 9); sheet.mergeCells(row, 10, row, 12); sheet.getCell(row, 1).value = `考核${tier.min}-${tier.max}`; sheet.getCell(row, 4).value = `每周${tier.fire}火、${tier.middle}中`; sheet.getCell(row, 10).value = `每周${tier.coins}币`; [1, 4, 10].forEach((col) => { sheet.getCell(row, col).alignment = center; sheet.getCell(row, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tier.color)}` } }; sheet.getCell(row, col).border = border }) })
  sheet.mergeCells('A24:L24'); sheet.getCell('A24').value = '特殊情况：扣包/或特殊情况奖励包'; sheet.getCell('A24').alignment = center; sheet.getCell('A24').border = border
  ranked.slice(0, 30).forEach((member, index) => { const row = index + 3; const rankCell = sheet.getCell(row, 13); rankCell.value = index + 1; rankCell.alignment = center; rankCell.border = border; rankCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tierForRank(index + 1).color)}` } }; [member.name, member.score ?? '', member.remark].forEach((value, i) => { const cell = sheet.getCell(row, 14 + i); cell.value = value; cell.alignment = center; cell.border = border; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorHex(tierForRank(index + 1).color)}` } } }) })
  powerRankMembers(members).slice(0, 30).forEach((member, index) => { const row = index + 3; const previousPower = member.previousPower ?? Math.max((member.power || 0) - (member.weeklyPower || 0), 0); [index + 1, member.name, previousPower, member.power || 0, member.weeklyPower || 0].forEach((value, i) => { const cell = sheet.getCell(row, 17 + i); cell.value = value; cell.alignment = center; cell.border = border }) })
  sheet.mergeCells('A26:L26'); sheet.getCell('A26').value = '战力、考核分排名每周六晚统计，每周更新'; sheet.getCell('A26').alignment = center
  return workbook
}

async function exportWorkbook(members: Member[], schedule: Schedule, ranked: Member[]) {
  const workbook = buildExportWorkbook(members, schedule, ranked)
  const buffer = await workbook.xlsx.writeBuffer(); const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); const url = URL.createObjectURL(blob); const link = document.createElement('a'); link.href = url; link.download = '繁星本周要塞包分配.xlsx'; link.click(); URL.revokeObjectURL(url)
}

const SHARED_STATE_ID = 'main'
type SharedState = { members: Member[]; queues: Record<AccessoryName, QueueEntry[]>; lastSweep: string; contest: boolean }
type SharedStateRow = { id: string; members: Member[]; queues: Record<AccessoryName, QueueEntry[]>; last_sweep: string; contest: boolean; updated_at?: string }

function sharedStateFingerprint(state: SharedState) {
  return JSON.stringify({ members: state.members, queues: state.queues, lastSweep: state.lastSweep, contest: state.contest })
}

function normalizeFuzzyText(value: string) {
  return value.trim().toLocaleLowerCase().replace(/\s+/g, '')
}

function fuzzyNameScore(name: string, query: string) {
  const text = normalizeFuzzyText(name)
  const term = normalizeFuzzyText(query)
  if (!term) return Number.POSITIVE_INFINITY
  if (text === term) return 0
  if (text.startsWith(term)) return 1
  const includedAt = text.indexOf(term)
  if (includedAt >= 0) return 10 + includedAt
  let cursor = 0
  let gaps = 0
  for (const character of term) {
    const matchedAt = text.indexOf(character, cursor)
    if (matchedAt < 0) return Number.POSITIVE_INFINITY
    gaps += matchedAt - cursor
    cursor = matchedAt + 1
  }
  return 100 + gaps
}

function fuzzyMemberNames(members: Member[], query: string, queuedEntries: QueueEntry[]) {
  const queuedNames = new Set(queuedEntries.map((entry) => normalizeFuzzyText(entry.name)))
  const seen = new Set<string>()
  return members
    .map((member, index) => ({ name: member.name.trim(), index, score: fuzzyNameScore(member.name, query) }))
    .filter(({ name, score }) => {
      const normalized = normalizeFuzzyText(name)
      if (!name || !Number.isFinite(score) || seen.has(normalized) || queuedNames.has(normalized)) return false
      seen.add(normalized)
      return true
    })
    .sort((a, b) => a.score - b.score || a.index - b.index)
    .slice(0, 6)
    .map(({ name }) => name)
}

function QueueNameInput({ accessory, value, members, queuedEntries, open, onOpen, onClose, onChange, onAdd }: { accessory: AccessoryName; value: string; members: Member[]; queuedEntries: QueueEntry[]; open: boolean; onOpen: () => void; onClose: () => void; onChange: (value: string) => void; onAdd: () => void }) {
  const suggestions = fuzzyMemberNames(members, value, queuedEntries)
  const listId = `queue-name-suggestions-${QUEUE_ACCESSORIES.findIndex((item) => item.name === accessory)}`
  return <div className="queue-add-shell" onBlur={(event) => { if (!event.currentTarget.contains(event.relatedTarget as Node | null)) onClose() }}>
    <div className="queue-add"><input role="combobox" aria-autocomplete="list" aria-expanded={open && Boolean(value.trim())} aria-controls={listId} value={value} onFocus={onOpen} onChange={(event) => { onChange(event.target.value); onOpen() }} onKeyDown={(event) => { if (event.key === 'Enter') onAdd() }} placeholder="输入成员姓名" /><button className="cat-add" onClick={onAdd}>＋</button></div>
    {open && value.trim() && <div className="queue-suggestions" id={listId} role="listbox" aria-label={`${accessory}成员姓名候选`}>
      {suggestions.length ? suggestions.map((name) => <button type="button" role="option" aria-selected="false" key={name} onClick={() => { onChange(name); onClose() }}>{name}</button>) : <span>未找到匹配成员，可继续手动输入</span>}
    </div>}
  </div>
}

export default function App() {
  const [members, setMembers] = useState<Member[]>(() => { try { const saved = localStorage.getItem('fortress-members'); return saved ? JSON.parse(saved).map((member: Member) => normalizeMember(member)) : [] } catch { return [] } })
  const [contest, setContest] = useState(false); const [notice, setNotice] = useState('已加载示例数据，可直接编辑或导入本周表格。'); const [activeSection, setActiveSection] = useState('matrix'); const fileRef = useRef<HTMLInputElement>(null)
  const undoStack = useRef<Member[][]>([])
  const redoStack = useRef<Member[][]>([])
  const [queues, setQueues] = useState<Record<AccessoryName, QueueEntry[]>>(() => { try { const saved = localStorage.getItem('fortress-accessory-queues'); return saved ? { ...emptyQueues(), ...JSON.parse(saved) } : emptyQueues() } catch { return emptyQueues() } })
  const [queueInputs, setQueueInputs] = useState<Record<AccessoryName, string>>(() => ({ 手镯: '', 戒指: '', 耳环: '', 腰带: '', 项链: '', 徽章: '', 剩余饰品: '' }))
  const [activeQueueInput, setActiveQueueInput] = useState<AccessoryName | null>(null)
  const [distributionCounts, setDistributionCounts] = useState<DistributionCounts>(() => emptyDistributionCounts())
  const [lastSweep, setLastSweep] = useState(() => localStorage.getItem('fortress-accessory-last-sweep') || '')
  const [cloudReady, setCloudReady] = useState(false)
  const hydratedRef = useRef(false)
  const syncEnabledRef = useRef(false)
  const lastSyncedFingerprintRef = useRef('')
  const lastRemoteUpdatedAtRef = useRef('')
  const saveSequenceRef = useRef(0)
  const realtimeStatusRef = useRef('')
  useEffect(() => {
    let cancelled = false
    const saveState = async (state: SharedState) => {
      if (!supabase) return
      const updatedAt = new Date().toISOString()
      const { error } = await supabase.from('fortress_state').upsert({ id: SHARED_STATE_ID, members: state.members, queues: state.queues, last_sweep: state.lastSweep, contest: state.contest, updated_at: updatedAt })
      if (error) throw new Error(error.message)
      lastRemoteUpdatedAtRef.current = updatedAt
      lastSyncedFingerprintRef.current = sharedStateFingerprint(state)
    }
    const loadSharedState = async () => {
      if (!supabase) { hydratedRef.current = true; setCloudReady(true); setNotice('当前未配置云端连接，数据只保存在本机。'); return }
      setNotice('正在读取共享数据…')
      try {
        const { data, error } = await supabase.from('fortress_state').select('members,queues,last_sweep,contest,updated_at').eq('id', SHARED_STATE_ID).maybeSingle()
        if (cancelled) return
        if (error) throw new Error(error.message)

        let nextMembers: Member[]
        let nextQueues: Record<AccessoryName, QueueEntry[]>
        let nextLastSweep: string
        let nextContest: boolean
        let noticeText: string
        if (data) {
          const row = data as SharedStateRow
          const remoteMembers = Array.isArray(row.members) ? row.members.map((member) => normalizeMember(member)) : []
          const shouldMigrateLocalRoster = members.length > remoteMembers.length
          const shouldSeedSample = !remoteMembers.length && !members.length
          nextMembers = shouldMigrateLocalRoster
            ? [...members, ...remoteMembers.filter((remote) => !members.some((local) => local.id === remote.id || local.name.trim() === remote.name.trim()))]
            : shouldSeedSample ? makeSampleMembers() : remoteMembers
          nextQueues = row.queues ? { ...emptyQueues(), ...row.queues } : emptyQueues()
          nextLastSweep = row.last_sweep || ''
          nextContest = Boolean(row.contest)
          noticeText = shouldMigrateLocalRoster ? '已将本机成员名单合并到共享数据。' : shouldSeedSample ? '没有找到成员数据，已自动导入以前的 30 人成员名单。' : '已连接共享数据，其他设备刷新后可看到最新内容。'
          lastRemoteUpdatedAtRef.current = row.updated_at || ''
          setMembers(nextMembers)
          setQueues(nextQueues)
          setLastSweep(nextLastSweep)
          setContest(nextContest)
          lastSyncedFingerprintRef.current = sharedStateFingerprint({ members: nextMembers, queues: nextQueues, lastSweep: nextLastSweep, contest: nextContest })
          if (shouldMigrateLocalRoster || shouldSeedSample) await saveState({ members: nextMembers, queues: nextQueues, lastSweep: nextLastSweep, contest: nextContest })
        } else {
          nextMembers = members.length ? members : makeSampleMembers()
          nextQueues = queues
          nextLastSweep = lastSweep
          nextContest = contest
          noticeText = members.length ? '已建立共享数据空间。' : '没有找到成员数据，已自动导入以前的 30 人成员名单。'
          setMembers(nextMembers)
          setQueues(nextQueues)
          setLastSweep(nextLastSweep)
          setContest(nextContest)
          await saveState({ members: nextMembers, queues: nextQueues, lastSweep: nextLastSweep, contest: nextContest })
        }
        syncEnabledRef.current = true
        if (!cancelled) setNotice(noticeText)
      } catch (error) {
        syncEnabledRef.current = false
        if (!cancelled) setNotice(`云端同步失败，暂时使用本机数据：${error instanceof Error ? error.message : '网络错误'}`)
      } finally {
        if (!cancelled) { hydratedRef.current = true; setCloudReady(true) }
      }
    }
    void loadSharedState()
    return () => { cancelled = true }
  }, [])
  useEffect(() => {
    if (!cloudReady || !supabase || !hydratedRef.current || !syncEnabledRef.current) return
    const state = { members, queues, lastSweep, contest }
    const fingerprint = sharedStateFingerprint(state)
    if (fingerprint === lastSyncedFingerprintRef.current) return
    const timer = window.setTimeout(() => {
      const sequence = ++saveSequenceRef.current
      const save = async () => {
        try {
          const updatedAt = new Date().toISOString()
          const { error } = await supabase.from('fortress_state').upsert({ id: SHARED_STATE_ID, members, queues, last_sweep: lastSweep, contest, updated_at: updatedAt })
          if (error) throw new Error(error.message)
          if (sequence === saveSequenceRef.current) {
            lastRemoteUpdatedAtRef.current = updatedAt
            lastSyncedFingerprintRef.current = fingerprint
            setNotice('已同步到云端，其他设备刷新后可看到最新内容。')
          }
        } catch (error) {
          if (sequence === saveSequenceRef.current) setNotice(`云端同步失败：${error instanceof Error ? error.message : '网络错误'}`)
        }
      }
      void save()
    }, 500)
    return () => window.clearTimeout(timer)
  }, [members, queues, lastSweep, contest, cloudReady])
  useEffect(() => {
    if (!cloudReady || !supabase || !syncEnabledRef.current) return
    const channel = supabase.channel(`fortress-state-${SHARED_STATE_ID}`)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'fortress_state', filter: `id=eq.${SHARED_STATE_ID}` }, (payload: { new?: SharedStateRow }) => {
        const row = payload.new
        if (!row || !Array.isArray(row.members)) return
        const nextMembers = row.members.map((member) => normalizeMember(member))
        const nextQueues = row.queues ? { ...emptyQueues(), ...row.queues } : emptyQueues()
        const nextLastSweep = row.last_sweep || ''
        const nextContest = Boolean(row.contest)
        const nextState = { members: nextMembers, queues: nextQueues, lastSweep: nextLastSweep, contest: nextContest }
        const fingerprint = sharedStateFingerprint(nextState)
        if (fingerprint === lastSyncedFingerprintRef.current) return
        lastRemoteUpdatedAtRef.current = row.updated_at || ''
        lastSyncedFingerprintRef.current = fingerprint
        setMembers(nextMembers)
        setQueues(nextQueues)
        setLastSweep(nextLastSweep)
        setContest(nextContest)
        setNotice('已接收其他设备的最新数据。')
      })
      .subscribe((status: string) => {
        if (status === 'SUBSCRIBED') {
          realtimeStatusRef.current = status
          return
        }
        if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT' || status === 'CLOSED') {
          if (realtimeStatusRef.current === status) return
          realtimeStatusRef.current = status
          setNotice('实时同步暂时断开，系统会自动重试并定时读取云端数据。')
        }
      })
    return () => { void supabase.removeChannel(channel) }
  }, [cloudReady])
  useEffect(() => {
    if (!cloudReady || !supabase || !syncEnabledRef.current) return
    let cancelled = false
    const pollRemoteState = async () => {
      try {
        const { data, error } = await supabase.from('fortress_state').select('members,queues,last_sweep,contest,updated_at').eq('id', SHARED_STATE_ID).maybeSingle()
        if (cancelled || error || !data) return
        const row = data as SharedStateRow
        const remoteUpdatedAt = Date.parse(row.updated_at || '')
        const currentUpdatedAt = Date.parse(lastRemoteUpdatedAtRef.current || '')
        const localState = { members, queues, lastSweep, contest }
        if (!Number.isFinite(remoteUpdatedAt) || remoteUpdatedAt <= currentUpdatedAt || sharedStateFingerprint(localState) !== lastSyncedFingerprintRef.current) return
        const nextMembers = Array.isArray(row.members) ? row.members.map((member) => normalizeMember(member)) : []
        const nextQueues = row.queues ? { ...emptyQueues(), ...row.queues } : emptyQueues()
        const nextState = { members: nextMembers, queues: nextQueues, lastSweep: row.last_sweep || '', contest: Boolean(row.contest) }
        const fingerprint = sharedStateFingerprint(nextState)
        if (fingerprint === lastSyncedFingerprintRef.current) return
        lastRemoteUpdatedAtRef.current = row.updated_at || ''
        lastSyncedFingerprintRef.current = fingerprint
        setMembers(nextMembers)
        setQueues(nextQueues)
        setLastSweep(nextState.lastSweep)
        setContest(nextState.contest)
        setNotice('已从云端读取其他设备的最新数据。')
      } catch {
        // Realtime remains the primary path; a later poll retries automatically.
      }
    }
    const interval = window.setInterval(() => { void pollRemoteState() }, 10000)
    const onFocus = () => { void pollRemoteState() }
    window.addEventListener('focus', onFocus)
    return () => { cancelled = true; window.clearInterval(interval); window.removeEventListener('focus', onFocus) }
  }, [cloudReady, members, queues, lastSweep, contest])
  const ranked = useMemo(() => rankMembers(members), [members]); const powerRanked = useMemo(() => powerRankMembers(members), [members]); const schedule = useMemo(() => buildAutoSchedule(ranked), [ranked]); const scoreMax = contest ? 57 : 37
  const weeklyPowerTotal = useMemo(() => members.reduce((total, member) => total + (member.weeklyPower || 0), 0), [members])
  const powerTotal = useMemo(() => members.reduce((total, member) => total + (member.power || 0), 0), [members])
  const previousPowerTotal = useMemo(() => members.reduce((total, member) => total + Math.max((member.power || 0) - (member.weeklyPower || 0), 0), 0), [members])
  const counts = useMemo(() => { const result = new Map<string, { fire: number; middle: number }>(); Object.entries(schedule).forEach(([key, id]) => { if (!id) return; const type = key.split(':')[1] as PackageType; const current = result.get(id) ?? { fire: 0, middle: 0 }; if (type === 'fire') current.fire += 1; else current.middle += 1; result.set(id, current) }); return result }, [schedule])
  useEffect(() => { localStorage.setItem('fortress-members', JSON.stringify(members)) }, [members])
  useEffect(() => { localStorage.setItem('fortress-accessory-queues', JSON.stringify(queues)) }, [queues])
  useEffect(() => {
    setDistributionCounts((current) => {
      let changed = false
      const next = { ...current }
      QUEUE_ACCESSORIES.forEach((accessory) => {
        const available = Math.min(2, queues[accessory.name].length)
        if (next[accessory.name] > available) { next[accessory.name] = available; changed = true }
      })
      return changed ? next : current
    })
  }, [queues])
  const commitMembers = (updater: (current: Member[]) => Member[], message?: string) => {
    setMembers((current) => {
      undoStack.current.push(cloneMembers(current))
      if (undoStack.current.length > 80) undoStack.current.shift()
      redoStack.current = []
      return updater(current).map((member) => normalizeMember(member))
    })
    if (message) setNotice(message)
  }
  const undoMembers = () => {
    const previous = undoStack.current.pop()
    if (!previous) { setNotice('没有可撤销的步骤。'); return }
    redoStack.current.push(cloneMembers(members))
    setMembers(cloneMembers(previous))
    setNotice('已撤销上一步。')
  }
  const redoMembers = () => {
    const next = redoStack.current.pop()
    if (!next) { setNotice('没有可前进的步骤。'); return }
    undoStack.current.push(cloneMembers(members))
    setMembers(cloneMembers(next))
    setNotice('已恢复下一步。')
  }
  const updateMember = (id: string, patch: Partial<Member>) => { commitMembers((current) => current.map((member) => member.id === id ? { ...member, ...patch } : member)) }
  const addMember = () => { const index = members.length + 1; commitMembers((current) => [...current, { id: `m-${Date.now()}`, name: `新成员${index}`, power: 0, previousPower: 0, weeklyPower: 0, score: null, remark: '', order: current.length }], '已新增成员，请填写姓名、战力和考核分。') }
  const removeMember = (id: string) => { commitMembers((current) => current.filter((member) => member.id !== id), '成员已删除，排名和矩阵已更新。') }
  const calculateWeeklyPower = () => { commitMembers((current) => current.map((member) => ({ ...member, previousPower: member.power, weeklyPower: 0 })), '已更新战力：本周战力已转为上周战力，提升已归零。') }
  const calculateGrowth = () => { commitMembers((current) => current.map((member) => ({ ...member, weeklyPower: Math.max((member.power || 0) - (member.previousPower || 0), 0) })), '已按“本周战力 - 上周战力”计算提升。') }
  const resetScores = () => { commitMembers((current) => current.map((member) => ({ ...member, score: 36 })), '考核分数已全部重置为 36。') }
  const visiblePowerMembers = powerRanked
  const visibleScoreMembers = ranked
  const clear = () => { commitMembers(() => [], '已清空成员数据。') }
  const reset = () => { commitMembers(() => makeSampleMembers(), '已恢复以前的 30 人成员数据。') }
  const addQueueEntry = (accessory: AccessoryName) => { const name = queueInputs[accessory].trim(); if (!name) { setNotice(`请先填写想要${accessory}的姓名。`); return } if (queues[accessory].some((entry) => entry.name.trim().toLowerCase() === name.toLowerCase())) { setNotice(`${name} 已经在${accessory}队列中。`); return } setQueues((current) => ({ ...current, [accessory]: [...current[accessory], { id: `q-${Date.now()}-${accessory}`, name, addedAt: new Date().toISOString() }] })); setQueueInputs((current) => ({ ...current, [accessory]: '' })); setNotice(`${name} 已加入${accessory}排队。`) }
  const removeQueueEntry = (accessory: AccessoryName, id: string) => { setQueues((current) => ({ ...current, [accessory]: current[accessory].filter((entry) => entry.id !== id) })); setNotice('已标记为分发完成，队列已更新。') }
  const setAccessoryDistributionCount = (accessory: AccessoryName, count: number) => {
    const available = Math.min(2, queues[accessory].length)
    setDistributionCounts((current) => ({ ...current, [accessory]: Math.max(0, Math.min(available, count)) }))
  }
  const sweepQueuesNow = () => {
    const actualCounts = QUEUE_ACCESSORIES.reduce((next, accessory) => {
      next[accessory.name] = Math.min(distributionCounts[accessory.name], queues[accessory.name].length)
      return next
    }, emptyDistributionCounts())
    const total = Object.values(actualCounts).reduce((sum, count) => sum + count, 0)
    if (!total) { setNotice('请先勾选至少一个有队首的部位，并设置发放数量。'); return }
    const distributed = QUEUE_ACCESSORIES.flatMap((accessory) => queues[accessory.name].slice(0, actualCounts[accessory.name]).map((entry) => `${accessory.name}（${entry.name}）`))
    setQueues((current) => sweepAccessoryQueues(current, actualCounts))
    setDistributionCounts(emptyDistributionCounts())
    setLastSweep(sundayDateKey())
    localStorage.setItem('fortress-accessory-last-sweep', sundayDateKey())
    setNotice(`已将 ${distributed.join('、')} 标记为已分发，共发放 ${total} 个队首名额。`)
  }
  const handleImport = async (event: React.ChangeEvent<HTMLInputElement>) => { const file = event.target.files?.[0]; if (!file) return; try { const imported = await importWorkbook(file); if (!imported.length) throw new Error('没有识别到成员'); const latest = new Map(members.map((member) => [member.name.trim(), member])); const prepared = imported.map((member) => { const saved = latest.get(member.name.trim()); return saved ? { ...member, id: saved.id, power: member.power, previousPower: saved.previousPower ?? saved.power, weeklyPower: 0, order: saved.order } : normalizeMember({ ...member, previousPower: member.power }) }); commitMembers(() => prepared, `已导入 ${prepared.length} 名成员；导入表中的战力已覆盖同名成员的旧数据，请按需点击“计算提升”。`) } catch (error) { setNotice(`导入失败：${error instanceof Error ? error.message : '文件格式不正确'}`) } finally { event.target.value = '' } }
  const goTo = (section: string) => { setActiveSection(section); document.getElementById(section)?.scrollIntoView({ behavior: 'smooth', block: 'start' }) }
  return <div className="app-shell">
    <span className="cat-sticker cat-art sticker-cat" style={{ backgroundImage: `url(${catStickerSheet})` }} aria-hidden="true"></span><span className="cat-sticker cat-art sticker-paw" style={{ backgroundImage: `url(${catStickerSheet})` }} aria-hidden="true"></span><span className="cat-sticker cat-art sticker-heart" style={{ backgroundImage: `url(${catStickerSheet})` }} aria-hidden="true"></span>
    <aside className="sidebar"><div className="brand-mark brand-cat" style={{ backgroundImage: `url(${catStickerSheet})` }}></div><h1>繁星要塞</h1><nav><button className={activeSection === 'matrix' ? 'active' : ''} onClick={() => goTo('matrix')}>▦ 要塞分包</button><button className={activeSection === 'members' ? 'active' : ''} onClick={() => goTo('members')}>♙ 成员管理与考核</button><button className={activeSection === 'accessories' ? 'active' : ''} onClick={() => goTo('accessories')}>◇ 饰品排队</button><button className={activeSection === 'instructions' ? 'active' : ''} onClick={() => goTo('instructions')}>▤ 使用说明</button></nav><div className="sidebar-note">{cloudReady && supabase ? <>数据已保存到共享云端<br />其他设备可同步</> : <>当前仅保存在本机<br />配置云端后可同步</>}</div></aside>
    <main className="content">
      <header className="topbar"><div><div className="eyebrow">FORTRESS DISTRIBUTION</div><h2>本周要塞包分配</h2></div><div className="toolbar"><button className="btn secondary" onClick={() => fileRef.current?.click()}>⇧ 导入 XLSX</button><input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleImport} /><button className="btn primary" onClick={() => exportWorkbook(members, schedule, ranked)}>⇩ 导出单表 XLSX</button></div></header>
      <section className="notice">{notice}<span className="notice-right"><label className="toggle"><input type="checkbox" checked={contest} onChange={(e) => setContest(e.target.checked)} /><span></span> 争霸周（最高 {scoreMax} 分）</label></span></section>
      <section className="stats"><div className="stat-card"><span>成员人数</span><strong>{members.length}</strong><small>{members.length === 30 ? '模板完整' : '目标 30 人'}</small></div><div className="stat-card"><span>当前最高分</span><strong>{ranked[0]?.score ?? '—'}</strong><small>普通周 37 · 争霸周 57</small></div><div className="stat-card"><span>本周提升</span><strong>{weeklyPowerTotal}</strong><small>点击右侧“更新战力”</small></div><div className="stat-card"><span>火/中包总量</span><strong>40 / 80</strong><small>8 个时段完整分配</small></div></section>
      <section id="matrix" className="panel matrix-panel"><div className="panel-heading"><div><span className="eyebrow">AUTO LAYOUT</span><h3>彩色分包矩阵</h3></div><button className="btn ghost" onClick={() => setNotice('矩阵会根据当前考核分排名自动生成。')}>↻ 重新生成矩阵</button></div><div className="matrix-scroll"><div className="matrix-grid">{SESSIONS.map((session) => <SessionBlock key={session.id} session={session} schedule={schedule} ranked={ranked} members={members} />)}</div></div><div className="legend">{TIERS.map((tier) => <span key={tier.min}><i className={`swatch ${tier.color}`}></i>{tier.min}-{tier.max} 名</span>)}</div></section>
      <section id="members" className="panel split-panel">
        <span id="ranking" className="anchor-target"></span>
        <div className="panel-heading">
          <div>
            <span className="eyebrow">MEMBER ROSTER</span>
            <h3>战力排行与考核分数</h3>
          </div>
          <div className="row-actions">
            <button className="btn ghost" onClick={undoMembers}>撤销</button>
            <button className="btn ghost" onClick={redoMembers}>前进</button>
            <button className="btn primary" onClick={addMember}>＋ 新增成员</button>
            <button className="btn ghost" onClick={reset}>恢复示例</button>
            <button className="btn danger" onClick={clear}>清空</button>
          </div>
        </div>
        <div className="split-tables">
          <div>
            <div className="table-heading">
              <div>
                <span className="eyebrow">POWER RANKING</span>
                <h4>战力排行表</h4>
              </div>
              <div className="table-heading-actions">
                <small>按本周战力排序</small>
                <button className="btn ghost" onClick={calculateGrowth}>计算提升</button>
                <button className="btn ghost" onClick={calculateWeeklyPower}>更新战力</button>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>战力排名</th>
                    <th>成员姓名</th>
                    <th>上周战力（w）</th>
                    <th>本周战力（w）</th>
                    <th>提升（w）</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {visiblePowerMembers.map((member) => {
                    const powerRank = powerRanked.findIndex((entry) => entry.id === member.id) + 1
                    const currentPower = member.power || 0
                    const previousPower = member.previousPower || 0
                    const growth = member.weeklyPower || 0
                    return (
                      <tr key={member.id}>
                        <td><span className="rank-pill">{powerRank || '—'}</span></td>
                        <td><input value={member.name} onChange={(e) => updateMember(member.id, { name: e.target.value })} /></td>
                        <td><input type="number" value={previousPower || ''} onChange={(e) => updateMember(member.id, { previousPower: e.target.value === '' ? 0 : Number(e.target.value) })} /></td>
                        <td><input type="number" value={currentPower || ''} onChange={(e) => updateMember(member.id, { power: e.target.value === '' ? 0 : Number(e.target.value) })} /></td>
                        <td><input type="number" value={growth || ''} onChange={(e) => updateMember(member.id, { weeklyPower: e.target.value === '' ? 0 : Number(e.target.value) })} /></td>
                        <td><button className="icon-btn" title="删除成员" onClick={() => removeMember(member.id)}>×</button></td>
                      </tr>
                    )
                  })}
                  <tr className="summary-row">
                    <td>合计</td>
                    <td>全部成员</td>
                    <td>{previousPowerTotal}</td>
                    <td>{powerTotal}</td>
                    <td>{weeklyPowerTotal}</td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
              {visiblePowerMembers.length === 0 && <div className="empty">没有匹配的成员，先添加一名试试。</div>}
            </div>
          </div>
          <div className="split-divider" aria-hidden="true">
            <span className="split-divider-sticker" style={{ backgroundImage: `url(${catStickerSheet})` }}></span>
            <span className="split-divider-line"></span>
            <span className="split-divider-sticker split-divider-sticker-alt" style={{ backgroundImage: `url(${catStickerSheet})` }}></span>
          </div>
          <div>
            <div className="table-heading">
              <div>
                <span className="eyebrow">SCORE RANKING</span>
                <h4>考核分数表</h4>
              </div>
              <div className="table-heading-actions">
                <small>分数相同按战力排序</small>
                <button className="btn ghost" onClick={resetScores}>重置分数</button>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>考核排名</th>
                    <th>成员姓名</th>
                    <th>考核分数</th>
                    <th>战力排名</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {visibleScoreMembers.map((member) => {
                    const scoreRank = ranked.findIndex((entry) => entry.id === member.id) + 1
                    const powerRank = powerRanked.findIndex((entry) => entry.id === member.id) + 1
                    return (
                      <tr key={member.id}>
                        <td><span className="rank-pill">{scoreRank || '—'}</span></td>
                        <td><input value={member.name} onChange={(e) => updateMember(member.id, { name: e.target.value })} /></td>
                        <td><input type="number" min="0" max={scoreMax} value={member.score || ''} placeholder="0" onChange={(e) => updateMember(member.id, { score: e.target.value === '' ? 0 : Number(e.target.value) })} /></td>
                        <td><span className={scoreRank ? `tier-dot ${tierForRank(scoreRank).color}` : 'rank-pill'}>{powerRank || '—'}</span></td>
                        <td><button className="icon-btn" title="删除成员" onClick={() => removeMember(member.id)}>×</button></td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
              {visibleScoreMembers.length === 0 && <div className="empty">没有匹配的成员，先添加一名试试。</div>}
            </div>
          </div>
        </div>
      </section>      <section id="instructions" className="panel rules-panel"><div><span className="eyebrow">PACKAGE RULES</span><h3>分包规则</h3></div><div className="protocol-copy"><p>每周六晚统计一次战力与考核分，按时参加活动基本不会扣包。</p><ul><li>考核分越高，分包排名越靠前；同分时战力高者优先。</li><li>普通周满分 37 分；争霸周满分 57 分。</li><li>每周固定 8 个时段，共 40 个火包和 80 个中包。</li><li>1–5 名领取 2 火 3 中，26–30 名领取 5 中，其余档位按卡片执行。</li><li>特殊奖励或扣包请在备注里写清楚。</li></ul></div><div className="tier-cards">{TIERS.map((tier) => <div className={`tier-card ${tier.color}`} key={tier.min}><b>{tier.min}-{tier.max}</b><span>{tier.fire} 火 · {tier.middle} 中</span><strong>{tier.coins} 币</strong></div>)}</div><p className="rule-copy">本规则以互相提醒、按时参加、公开透明为原则。</p></section>
      <section id="accessories" className="panel accessory-panel">
        <div className="panel-heading">
          <div><span className="eyebrow">ACCESSORY QUEUE</span><h3>饰品排队</h3></div>
          <button className="btn ghost" onClick={sweepQueuesNow} disabled={!Object.values(distributionCounts).some(Boolean)}>↻ 发放首位{Object.values(distributionCounts).some(Boolean) ? `（${Object.values(distributionCounts).reduce((sum, count) => sum + count, 0)}）` : ''}</button>
        </div>
        <p className="accessory-intro">选择需要的饰品并留下姓名。发放时可多选需要的部位（包含剩余饰品），每个部位可发放 1 或 2 个队首名额。</p>
        <fieldset className="accessory-picker">
          <legend>勾选部位并设置发放数量（0–2）</legend>
          {QUEUE_ACCESSORIES.map((accessory) => {
            const first = queues[accessory.name][0]
            const count = distributionCounts[accessory.name]
            const available = Math.min(2, queues[accessory.name].length)
            const waitingNames = queues[accessory.name].slice(0, 2).map((entry) => entry.name).join('、')
            return <div className={`accessory-choice ${count ? 'selected' : ''} ${!first ? 'disabled' : ''}`} key={accessory.name}>
              <label><input type="checkbox" checked={count > 0} disabled={!first} onChange={(event) => setAccessoryDistributionCount(accessory.name, event.target.checked ? 1 : 0)} /><span>{accessory.icon} {accessory.name}</span></label>
              <small>{first ? `${queues[accessory.name].length > 1 ? '前两位' : '队首'}：${waitingNames}` : '暂无队首'}</small>
              <div className="quantity-stepper" aria-label={`${accessory.name}发放数量`}>
                <button type="button" disabled={count === 0} onClick={() => setAccessoryDistributionCount(accessory.name, count - 1)} aria-label={`减少${accessory.name}发放数量`}>−</button>
                <strong>{count}</strong>
                <button type="button" disabled={count >= available} onClick={() => setAccessoryDistributionCount(accessory.name, count + 1)} aria-label={`增加${accessory.name}发放数量`}>＋</button>
              </div>
            </div>
          })}
        </fieldset>
        <div className="accessory-grid">
          {QUEUE_ACCESSORIES.map((accessory) => <div className={`accessory-card ${accessory.color}`} key={accessory.name}><div className="accessory-title"><span className="accessory-icon">{accessory.icon}</span><div><strong>{accessory.name}</strong><small>{accessory.name === '剩余饰品' ? '任意无人认领的部位' : `${queues[accessory.name].length} 人排队`}</small>{accessory.name === '剩余饰品' && <small>{queues[accessory.name].length} 人排队</small>}</div></div><QueueNameInput accessory={accessory.name} value={queueInputs[accessory.name]} members={members} queuedEntries={queues[accessory.name]} open={activeQueueInput === accessory.name} onOpen={() => setActiveQueueInput(accessory.name)} onClose={() => setActiveQueueInput(null)} onChange={(value) => setQueueInputs((current) => ({ ...current, [accessory.name]: value }))} onAdd={() => { addQueueEntry(accessory.name); setActiveQueueInput(null) }} />{queues[accessory.name].length ? <ol className="queue-list">{queues[accessory.name].map((entry, index) => <li key={entry.id}><span className="queue-number">{index + 1}</span><span className="queue-name">{entry.name}</span><button onClick={() => removeQueueEntry(accessory.name, entry.id)}>已分发</button></li>)}</ol> : <div className="queue-empty">暂无排队</div>}</div>)}
          {[1, 2].map((slot) => <div className="accessory-card development-card" key={`development-${slot}`}><div className="development-placeholder"><span>🚧</span><strong>开发中</strong></div></div>)}
        </div>
        <p className="queue-note">发放不会自动执行；每个部位（包含剩余饰品）只能选择 0、1、2，系统会按数量依次发放队首。队列数据会保存在共享数据中。</p>
      </section>
      <footer className="signature">署名：繁星</footer>
    </main>
  </div>
}

function SessionBlock({ session, schedule, ranked, members }: { session: Session; schedule: Schedule; ranked: Member[]; members: Member[] }) {
  const memberById = new Map(members.map((member) => [member.id, member]))
  return <div className={`session-block ${session.block}`}><div className="session-title">🐾 {session.label}</div><div className="session-types">{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, 0)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <div className={rank ? `header-fill ${tierForRank(rank).color}` : ''} key={type}>{PACKAGE_LABELS[type]}</div> })}</div>{Array.from({ length: 5 }, (_, row) => <div className="session-row" key={row}>{TYPE_ORDER.map((type) => { const id = schedule[cellKey(session.id, type, row)] ?? ''; const member = memberById.get(id); const rank = member ? ranked.findIndex((entry) => entry.id === member.id) + 1 : 0; return <div key={type} className={`matrix-name ${rank ? `rank-fill ${tierForRank(rank).color}` : ''}`}>{member?.name ?? '—'}</div> })}</div>)}</div>
}

