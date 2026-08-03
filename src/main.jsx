import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import './styles.css'

const dataFiles = import.meta.glob('../data/*/*', { eager: true, as: 'url' })
const matchdayFiles = import.meta.glob('../data/*/*/**/*', { eager: true, as: 'url' })
const bonusIcons = import.meta.glob('../icone/bonus_*.png', { eager: true, as: 'url' })
const groupedData = Object.entries(dataFiles).reduce((result, [path, url]) => {
  const match = path.match(/data\/([^/]+)\/([^/]+)$/)
  if (!match) return result
  const [, year, filename] = match
  const type = filename.replace(/\.(csv|xls)$/i, '')
  result[year] ||= {}
  result[year][type] = url
  return result
}, {})
const availableSeasons = Object.keys(groupedData).sort()
const groupedMatchdays = Object.entries(matchdayFiles).reduce((result, [path, url]) => {
  const match = path.match(/data\/([^/]+)\/(campionato|fantapunti)\/([^/]+)\/(.+)$/)
  if (!match) return result
  const [, season, competition, day, relativePath] = match
  result[season] ||= {}
  result[season][competition] ||= {}
  result[season][competition][day] ||= { total: '', teams: {} }
  if (relativePath === 'totale.csv') result[season][competition][day].total = url
  else if (relativePath.endsWith('/giocatori.csv')) result[season][competition][day].teams[relativePath.replace(/\/giocatori\.csv$/, '')] = url
  return result
}, {})
const groupedKnockout = Object.entries(matchdayFiles).reduce((result, [path, url]) => {
  const match = path.match(/data\/([^/]+)\/(champions|europa_league|coppa_italia)\/([^/]+)\/(.+)$/)
  if (!match) return result
  const [, season, competition, stage, remainder] = match
  const parts = remainder.split('/')
  const hasDayFolder = (parts.length === 2 && parts[1] === 'totale.csv') || (parts.length === 3 && parts[2] === 'giocatori.csv')
  const day = hasDayFolder ? parts.shift() : ''
  const relativePath = parts.join('/')
  if (relativePath !== 'totale.csv' && !relativePath.endsWith('/giocatori.csv')) return result
  result[season] ||= {}
  result[season][competition] ||= {}
  const key = day ? `${stage}/${day}` : stage
  result[season][competition][key] ||= { total: '', teams: {} }
  if (relativePath === 'totale.csv') result[season][competition][key].total = url
  else if (relativePath.endsWith('/giocatori.csv')) result[season][competition][key].teams[relativePath.slice(0, -'/giocatori.csv'.length)] = url
  return result
}, {})

const navItems = [
  { id: 'fantasy', label: 'Fantapunti', icon: '✦' },
  { id: 'championship', label: 'Campionato', icon: '⌁' },
  { id: 'matchdays', label: 'Giornate', icon: '◷' },
  { id: 'champions', label: 'Champions', icon: '◆' },
  { id: 'scorers', label: 'Capocannonieri', icon: '⚽' },
  { id: 'users', label: 'Utenti', icon: '◎' },
  { id: 'rosters', label: 'Rose', icon: '♟' },
]

function parseCsv(text) {
  const lines = text.replace(/^\uFEFF/, '').trim().split(/\r?\n/)
  const split = (line) => {
    const cells = []
    let cell = ''
    let quoted = false
    for (let index = 0; index < line.length; index += 1) {
      const character = line[index]
      if (character === '"') {
        if (quoted && line[index + 1] === '"') {
          cell += '"'
          index += 1
        } else {
          quoted = !quoted
        }
      } else if (character === ';' && !quoted) {
        cells.push(cell.trim())
        cell = ''
      } else {
        cell += character
      }
    }
    cells.push(cell.trim())
    return cells
  }
  const headers = split(lines[0])
  return lines.slice(1).filter(Boolean).map((line) => Object.fromEntries(split(line).map((value, i) => [headers[i], value])))
}

function numeric(value) {
  return Number(String(value ?? '').replace(',', '.')) || 0
}

function rankingValue(row, key) {
  if (key === 'PG') return row.PG ? numeric(row.PG) : numeric(row.Vittorie) + numeric(row.Nulle) + numeric(row.Sconfitte)
  return numeric(row[key])
}

function normalizeTeam(value) {
  return String(value ?? '').trim().toLocaleLowerCase('it-IT')
}

function changeSeason(event, setSeason, setPage) {
  setPage('championship')
  setSeason(event.target.value)
}

function withUsers(rows, users) {
  const owners = new Map()
  for (const row of users) {
    const team = normalizeTeam(row.Squadra)
    const user = String(row.Utente || '').trim()
    if (!team || !user) continue
    if (!owners.has(team)) owners.set(team, [])
    if (!owners.get(team).includes(user)) owners.get(team).push(user)
  }
  return rows.map((row) => {
    const teamOwners = owners.get(normalizeTeam(row.Squadra)) || []
    return { ...row, Utente: teamOwners.join(' / ') || '—', Utenti: teamOwners }
  })
}

async function fetchCupWinners(season, users) {
  const competitions = [
    ['champions', 'Champions'],
    ['europa_league', 'Europa League'],
    ['coppa_italia', 'Coppa Italia'],
  ]
  return Promise.all(competitions.map(async ([slug, label]) => {
    const stages = groupedKnockout[season]?.[slug] || {}
    const stageRows = await Promise.all(Object.entries(stages).map(async ([stage, files]) => {
      if (!files.total) return [stage, []]
      const response = await fetch(files.total)
      return [stage, response.ok ? parseCsv(await response.text()) : []]
    }))
    const rowsByStage = Object.fromEntries(stageRows)
    const finalRows = rowsByStage.finale || []
    const teamsIn = (rows) => new Set(rows.flatMap((row) => [row.Casa, row.Trasferta]).filter(Boolean).map(normalizeTeam))
    const finalTeams = teamsIn(finalRows)
    const semiTeams = new Set(stageRows.filter(([stage]) => /semi/i.test(stage)).flatMap(([, rows]) => [...teamsIn(rows)]))
    const quarterTeams = new Set(stageRows.filter(([stage]) => /quart/i.test(stage)).flatMap(([, rows]) => [...teamsIn(rows)]))
    const positions = []
    const addPosition = (team, position, stage) => {
      const teamUsers = users.filter((row) => normalizeTeam(row.Squadra) === normalizeTeam(team)).map((row) => row.Utente).filter(Boolean)
      positions.push({ team, position, stage, users: [...new Set(teamUsers)] })
    }
    finalRows.slice(0, 2).forEach((match, matchIndex) => {
      const result = String(match.Risultato || '').match(/(\d+)\s*-\s*(\d+)/)
      if (!result) return
      const homeWon = Number(result[1]) > Number(result[2])
      const awayWon = Number(result[2]) > Number(result[1])
      if (!homeWon && !awayWon) return
      const firstPosition = matchIndex === 0 ? 1 : 3
      const secondPosition = firstPosition + 1
      addPosition(match.Casa, homeWon ? firstPosition : secondPosition, 'finale')
      addPosition(match.Trasferta, awayWon ? firstPosition : secondPosition, 'finale')
    })
    for (const team of semiTeams) if (!finalTeams.has(team)) {
      const original = [...stageRows.flatMap(([, rows]) => rows.flatMap((row) => [row.Casa, row.Trasferta]))].find((candidate) => normalizeTeam(candidate) === team)
      addPosition(original, null, 'semifinali')
    }
    for (const team of quarterTeams) if (!semiTeams.has(team) && !finalTeams.has(team)) {
      const original = [...stageRows.flatMap(([, rows]) => rows.flatMap((row) => [row.Casa, row.Trasferta]))].find((candidate) => normalizeTeam(candidate) === team)
      addPosition(original, null, 'quarti')
    }
    return { slug, label, positions }
  }))
}

async function fetchSeasonSummary(season) {
  const files = groupedData[season] || {}
  const types = ['utenti', 'fantapunti', 'campionato', 'capocannoniere']
  const entries = await Promise.all(types.map(async (type) => {
    if (!files[type]) return [type, '']
    const response = await fetch(files[type])
    return [type, response.ok ? await response.text() : '']
  }))
  const contents = Object.fromEntries(entries)
  const championsFile = files.champions || files.champions_gironi
  const championsText = championsFile ? await (await fetch(championsFile)).text() : ''
  const users = contents.utenti ? parseCsv(contents.utenti) : []
  const cups = await fetchCupWinners(season, users)
  return {
    season,
    users,
    fantasy: withUsers(contents.fantapunti ? parseCsv(contents.fantapunti) : [], users),
    championship: withUsers(contents.campionato ? parseCsv(contents.campionato) : [], users),
    champions: withUsers(championsText ? parseCsv(championsText) : [], users),
    scorers: withUsers(contents.capocannoniere ? parseCsv(contents.capocannoniere) : [], users),
    cups,
  }
}

function App() {
  const [page, setPage] = useState('championship')
  const [expandedCompetitions, setExpandedCompetitions] = useState({ championship: true, fantasy: false })
  const [season, setSeason] = useState(availableSeasons[availableSeasons.length - 1] || '')
  const seasons = availableSeasons
  const [data, setData] = useState({ fantasy: [], championship: [], champions: [], scorers: [], users: [], rosters: [] })
  const [seasonHistory, setSeasonHistory] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')

  useEffect(() => {
    async function load() {
      setLoading(true)
      setError('')
      try {
        if (!availableSeasons.length) throw new Error('Nessun dato trovato nella cartella data')
        const files = groupedData[season]
        const responses = await Promise.all(Object.entries(files).map(async ([type, url]) => [type, await fetch(url)]))
        if (responses.some(([, response]) => !response.ok)) throw new Error('Impossibile leggere i file dati')
        const contents = Object.fromEntries(await Promise.all(responses.map(async ([type, response]) => [type, await (type === 'rose' ? response.arrayBuffer() : response.text())])))
        const fantasyText = contents.fantapunti
        const scorersText = contents.capocannoniere
        const championshipText = contents.campionato
        const championsText = contents.champions || contents.champions_gironi
        const users = contents.utenti ? parseCsv(contents.utenti) : []
        const rosterBuffer = contents.rose
        if (!fantasyText) throw new Error(`File fantapunti mancante nella stagione ${season}`)
        // SheetJS is more reliable in the browser when given a typed byte array.
        const workbook = rosterBuffer ? XLSX.read(new Uint8Array(rosterBuffer), { type: 'array' }) : null
        const rosters = workbook ? workbook.SheetNames.flatMap((sheetName) => {
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' })
          const header = rows[0] || []
          const teams = []
          for (let column = 0; column < header.length; column += 4) {
            if (header[column + 1]) teams.push({ column, name: String(header[column + 1]).trim() })
          }
          return teams.flatMap(({ column, name }) => rows.slice(1).map((row) => ({
            squadra: name,
            ruolo: row[column] || '',
            giocatore: row[column + 1] || '',
            quotazione: row[column + 2] || '',
            club: row[column + 3] || '',
            sheet: sheetName,
          })).filter((row) => row.giocatore))
        }) : []
        setData({ fantasy: withUsers(parseCsv(fantasyText), users), championship: withUsers(championshipText ? parseCsv(championshipText) : [], users), champions: withUsers(championsText ? parseCsv(championsText) : [], users), scorers: withUsers(scorersText ? parseCsv(scorersText) : [], users), users, rosters })
        setSeasonHistory(await Promise.all(seasons.map(fetchSeasonSummary)))
      } catch (err) {
        setError(err.message)
      } finally { setLoading(false) }
    }
    load()
  }, [season])

  useEffect(() => {
    if (loading) return
    const currentNav = navItems.find((item) => item.id === page)
    if (currentNav?.id === 'championship' && !data.championship.length) setPage('fantasy')
    if (currentNav?.id === 'champions' && !data.champions.length) setPage('fantasy')
  }, [data.championship, data.champions, loading, page])

  const teams = useMemo(() => [...new Set(data.rosters.map((row) => row.squadra).filter(Boolean))], [data.rosters])
  const users = useMemo(() => {
    const allUsers = seasonHistory.flatMap((item) => item.users.map((row) => row.Utente)).concat(data.users.map((row) => row.Utente))
    return [...new Set(allUsers.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'it'))
  }, [data.users, seasonHistory])
  const current = navItems.find((item) => item.id === page)
  const matchdayData = groupedMatchdays[season] || {}
  const visibleNavItems = navItems.filter((item) => (item.id !== 'championship' || data.championship.length) && (item.id !== 'champions' || data.champions.length))
  const pageLabel = page === 'matchdays-campionato' ? 'Giornate di campionato' : page === 'matchdays-fantapunti' ? 'Giornate di Fantapunti' : page === 'champions-knockout' ? 'Fase Eliminazione diretta' : current?.label || 'Fantapunti'
  const seasonLabel = season.replace('_', '/')

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><span className="brand-mark">F</span><span>Fantapazz<br /><small>STATS</small></span></div>
      <button className={`global-user-link ${page === 'users' ? 'active' : ''}`} onClick={() => setPage('users')}><span>◎</span>Utenti</button>
      <label className="season"><span className="live-dot" /> Stagione<select value={season} onChange={(event) => changeSeason(event, setSeason, setPage)}>{seasons.map((value) => <option key={value} value={value}>{value.replace('_', '/')}</option>)}</select></label>
      <nav>{visibleNavItems.filter((item) => item.id !== 'users' && item.id !== 'matchdays').map((item) => { const dataKey = item.id === 'championship' ? 'campionato' : item.id === 'fantasy' ? 'fantapunti' : ''; const hasMatchdaySubmenu = (item.id === 'championship' || item.id === 'fantasy') && matchdayData[dataKey]; const hasChampionsSubmenu = item.id === 'champions' && data.champions.length; const hasKnockoutSubmenu = item.id === 'champions' && groupedKnockout[season]?.champions; const hasSubmenu = hasMatchdaySubmenu || hasChampionsSubmenu; const submenuPage = item.id === 'championship' ? 'matchdays-campionato' : item.id === 'fantasy' ? 'matchdays-fantapunti' : 'champions-knockout'; const submenuLabel = item.id === 'champions' ? 'Fase Eliminazione diretta' : 'Giornate'; const isSubmenuActive = page === submenuPage; return <React.Fragment key={item.id}><button className={page === item.id || isSubmenuActive ? 'active' : ''} onClick={() => { setPage(item.id); if (hasSubmenu) setExpandedCompetitions((current) => ({ ...current, [item.id]: !current[item.id] })) }}><span>{item.icon}</span>{item.label}</button>{hasSubmenu && expandedCompetitions[item.id] && <>{item.id === 'champions' && <button className={`nav-subitem ${page === 'champions' ? 'active' : ''}`} onClick={() => setPage('champions')}><span>◆</span>Gironi</button>}{(item.id !== 'champions' || hasKnockoutSubmenu) && <button className={`nav-subitem ${isSubmenuActive ? 'active' : ''}`} onClick={() => setPage(submenuPage)}><span>{item.id === 'champions' ? '◇' : '◷'}</span>{submenuLabel}</button>}</>}</React.Fragment> })}</nav>
      <div className="sidebar-foot">DATI LOCALI<br /><span>Importati dai file CSV</span></div>
    </aside>
    <main className="content">
      <header className="topbar"><div className="mobile-brand"><span className="brand-mark">F</span> Fantapazz Stats</div><div className="top-season">{seasonLabel} <span>•</span> Lega Fantapazz</div><label className="mobile-season"><span>Stagione</span><select aria-label="Seleziona stagione" value={season} onChange={(event) => changeSeason(event, setSeason, setPage)}>{seasons.map((value) => <option key={value} value={value}>{value.replace('_', '/')}</option>)}</select></label><div className="avatar">FP</div></header>
      <section className="page-heading"><div><p className="eyebrow">PANORAMICA DELLA LEGA</p><h1>{pageLabel}</h1><p className="subheading">Classifica e statistiche della stagione {seasonLabel}</p></div><div className="data-badge"><span className="pulse" /> Dati aggiornati</div></section>
      {loading && <div className="state-card">Caricamento dati…</div>}
      {error && <div className="state-card error">{error}. Avvia l’app tramite <code>npm run dev</code> per servire i file locali.</div>}
      {!loading && !error && page === 'fantasy' && <Ranking rows={data.fantasy} type="fantasy" />}
      {!loading && !error && page === 'championship' && <Ranking rows={data.championship} type="championship" />}
      {!loading && !error && page === 'matchdays-campionato' && <Matchdays days={matchdayData.campionato || {}} competition="campionato" />}
      {!loading && !error && page === 'matchdays-fantapunti' && <Matchdays days={matchdayData.fantapunti || {}} competition="fantapunti" />}
      {!loading && !error && page === 'champions' && <ChampionsRanking rows={data.champions} />}
      {!loading && !error && page === 'champions-knockout' && <Matchdays days={groupedKnockout[season]?.champions || {}} competition="champions" />}
      {!loading && !error && page === 'scorers' && <Ranking rows={data.scorers} type="scorers" />}
      {!loading && !error && page === 'users' && <UserStats history={seasonHistory} users={users} />}
      {!loading && !error && page === 'rosters' && <Roster rows={data.rosters} teams={teams} users={data.users} />}
    </main>
  </div>
}

function Ranking({ rows, type }) {
  const isChampionship = type === 'championship'
  const isChampions = type === 'champions'
  const isScorers = type === 'scorers'
  const scoreLabel = isChampionship || isChampions ? 'Punti' : 'FantaPunti'
  const title = isChampionship ? 'Classifica campionato' : isChampions ? 'Classifica Champions' : isScorers ? 'Classifica capocannonieri' : 'Classifica fantapunti'
  const isLeague = isChampionship || isChampions
  const bestValues = (isLeague ? ['Punti', 'PG', 'Vittorie', 'Nulle', 'Sconfitte', 'GolFa', 'GolSu', 'FantaPunti'] : ['FantaPunti']).reduce((result, key) => {
    result[key] = Math.max(...rows.map((row) => rankingValue(row, key)))
    return result
  }, {})
  const bestClass = (row, key) => rankingValue(row, key) === bestValues[key] ? 'best-value' : ''
  return <div className="panel"><div className="panel-toolbar"><div><h2>{title}</h2><p>{rows.length} squadre partecipanti</p></div><div className="count-pill">{rows.length} <span>squadre</span></div></div><div className="table-wrap"><table className={isLeague ? 'league-table' : 'compact-table'}><thead><tr><th>#</th><th>Squadra</th><th>Utente</th>{isLeague && <><th className="right">Punti</th><th>PG</th><th>V</th><th>N</th><th>P</th><th>GF</th><th>GS</th><th className="right">FantaPunti</th></>} {!isLeague && <th className="right">{scoreLabel}</th>}</tr></thead><tbody>{rows.map((row, index) => <tr key={`${row.Squadra}-${index}`}><td><Rank rank={index + 1} /></td><td className="team-name">{row.Squadra}</td><td className="owner">{row.Utente}</td>{isLeague && <><td className={`score ${bestClass(row, 'Punti')}`}>{numeric(row.Punti).toLocaleString('it-IT')}</td><td className={bestClass(row, 'PG')}>{row.PG || (numeric(row.Vittorie) + numeric(row.Nulle) + numeric(row.Sconfitte))}</td><td className={bestClass(row, 'Vittorie')}>{row.Vittorie}</td><td className={bestClass(row, 'Nulle')}>{row.Nulle}</td><td className={bestClass(row, 'Sconfitte')}>{row.Sconfitte}</td><td className={bestClass(row, 'GolFa')}>{row.GolFa}</td><td className={bestClass(row, 'GolSu')}>{row.GolSu}</td><td className={`score ${bestClass(row, 'FantaPunti')}`}>{numeric(row.FantaPunti).toLocaleString('it-IT', { minimumFractionDigits: String(row.FantaPunti).includes(',') ? 1 : 0 })}</td></>}{!isLeague && <td className={`score ${bestClass(row, 'FantaPunti')}`}>{numeric(row.FantaPunti).toLocaleString('it-IT', { minimumFractionDigits: String(row.FantaPunti).includes(',') ? 1 : 0 })}</td>}</tr>)}</tbody></table></div></div>
}

function ChampionsRanking({ rows }) {
  const groups = [...new Set(rows.map((row) => row.Girone).filter(Boolean))]
  return <div className="champions-groups">{groups.map((group) => <div key={group} className="champions-group"><div className="group-heading"><span>Champions</span><strong>Girone {group}</strong></div><Ranking rows={rows.filter((row) => row.Girone === group)} type="champions" /></div>)}</div>
}

function Rank({ rank }) { return rank <= 3 ? <span className={`medal medal-${rank}`}>{rank}</span> : <span className="rank">{rank}</span> }
function HistoryResult({ rank, value, prefix = '' }) { return <span className="history-result">{prefix}<Rank rank={rank} /><span>{value}</span></span> }
function CupResult({ result }) { return result?.position ? <Rank rank={result.position} /> : result?.stage ? <span className="cup-stage">{result.stage === 'semifinali' ? 'Semi' : 'Quarti'}</span> : '-' }

function bonusIconUrl(code) {
  const match = String(code).match(/imgBonus_(\d+)/i)
  if (!match) return ''
  const path = Object.keys(bonusIcons).find((value) => value.endsWith(`/bonus_${match[1]}.png`))
  return path ? bonusIcons[path] : ''
}

function bonusCodes(value) {
  return [...String(value || '').matchAll(/imgBonus_(\d+)/gi)].map((match) => match[1])
}

function voteClass(value) {
  if (value === '') return 'vote-na'
  const vote = numeric(value)
  return vote > 6 ? 'vote-high' : vote < 6 ? 'vote-low' : 'vote-neutral'
}

function TeamBonusList({ bonuses }) {
  const labels = [['Difesa', bonuses?.Difesa], ['Centrocampo', bonuses?.Centrocampo], ['Attacco', bonuses?.Attacco]].filter(([, value]) => String(value || '').trim())
  if (!labels.length) return null
  return <div className="team-bonuses">{labels.map(([label, value]) => <span className={numeric(value) >= 0 ? 'team-bonus-positive' : 'team-bonus-negative'} key={label}><small>{label}</small><strong>{value}</strong></span>)}</div>
}

function PlayerMatchList({ team, rows, side, bonuses }) {
  return <section className={`match-team match-team-${side}`}><h3>{team}</h3><div className="match-players">{rows.length ? rows.map((player, index) => { const role = String(player.Ruolo || '').trim().toLowerCase().charAt(0); const roleClass = { p: 'por', d: 'dif', c: 'cen', a: 'att' }[role] || role; const codes = bonusCodes(player.BonusMalus); const roleElement = <span className={`role role-${roleClass}`}>{player.Ruolo || '—'}</span>; const nameElement = <span className="match-player-name">{player.Giocatore}</span>; const voteElement = <span className={`match-vote ${voteClass(String(player.Voto || '').trim())}`}>{player.Voto || '—'}</span>; const bonusElement = <span className={`bonus-icons ${codes.length === 0 ? 'bonus-icons-empty' : ''}`}>{codes.map((code, bonusIndex) => { const src = bonusIconUrl(`imgBonus_${code}`); return src ? <img key={`${code}-${bonusIndex}`} src={src} alt={`Bonus ${code}`} title={`Bonus ${code}`} /> : <span key={`${code}-${bonusIndex}`} title={`Bonus ${code}`}>+{code}</span> })}</span>; return <div className={`match-player ${player.Stato === 'Riserva' ? 'reserve' : ''}`} key={`${player.Giocatore}-${index}`}>{side === 'home' ? <>{roleElement}{nameElement}{bonusElement}{voteElement}</> : <>{voteElement}{bonusElement}{nameElement}{roleElement}</>}</div> }) : <p className="empty">Nessun giocatore trovato.</p>}</div><TeamBonusList bonuses={bonuses} /></section>
}

function Matchdays({ days, competition }) {
  const isFantasy = competition === 'fantapunti'
  const dayNames = Object.keys(days).sort((a, b) => Number.parseInt(a, 10) - Number.parseInt(b, 10))
  const [selectedDay, setSelectedDay] = useState(dayNames[0] || '')
  const [dayData, setDayData] = useState({ matches: [], players: {}, loading: false, error: '' })
  useEffect(() => setSelectedDay(dayNames[0] || ''), [days])
  useEffect(() => {
    let cancelled = false
    async function loadDay() {
      if (!selectedDay || !days[selectedDay]?.total) return
      setDayData((current) => ({ ...current, loading: true, error: '' }))
      try {
        const day = days[selectedDay]
        const totalResponse = await fetch(day.total)
        if (!totalResponse.ok) throw new Error('Impossibile leggere il totale della giornata')
        const matches = parseCsv(await totalResponse.text())
        const teams = [...new Set(isFantasy ? matches.map((match) => match.Squadra).filter(Boolean) : matches.flatMap((match) => [match.Casa, match.Trasferta]).filter(Boolean))]
        const playerEntries = await Promise.all(teams.map(async (team) => {
          const url = day.teams[team]
          if (!url) return [team, []]
          const response = await fetch(url)
          return [team, response.ok ? parseCsv(await response.text()) : []]
        }))
        if (!cancelled) setDayData({ matches, players: Object.fromEntries(playerEntries), loading: false, error: '' })
      } catch (error) {
        if (!cancelled) setDayData({ matches: [], players: {}, loading: false, error: error.message })
      }
    }
    loadDay()
    return () => { cancelled = true }
  }, [days, isFantasy, selectedDay])
  const title = isFantasy ? 'Giornate di Fantapunti' : competition === 'champions' ? 'Fase Eliminazione diretta' : 'Giornate di campionato'
  const formatDay = (day) => day.split('/').map((part) => part.replaceAll('_', ' ')).join(' · ')
  return <div className="matchdays-page"><div className="panel matchday-selector"><div className="panel-toolbar"><div><h2>{title}</h2><p>Visualizza i punteggi e le formazioni della giornata</p></div><label className="matchday-inline-selector">Seleziona giornata<select value={selectedDay} onChange={(event) => setSelectedDay(event.target.value)}>{dayNames.map((day) => <option key={day} value={day}>{formatDay(day)}</option>)}</select></label></div></div>{dayData.loading && <div className="state-card">Caricamento giornata…</div>}{dayData.error && <div className="state-card error">{dayData.error}</div>}{!dayData.loading && !dayData.error && (isFantasy ? <div className="panel fantasy-ranking-day"><div className="panel-toolbar"><div><h2>Classifica Fantapunti</h2><p>Risultati della giornata</p></div></div><div className="table-wrap"><table className="compact-table"><thead><tr><th>#</th><th>Squadra</th><th className="right">FantaPunti</th></tr></thead><tbody>{dayData.matches.map((match, index) => <tr key={`${match.Squadra}-${index}`}><td><Rank rank={index + 1} /></td><td className="team-name">{match.Squadra}</td><td className="score">{match.FantaPunti}</td></tr>)}</tbody></table></div></div> : <div className="matchday-list">{dayData.matches.map((match, index) => <article className="match-card" key={`${match.Casa}-${match.Trasferta}-${index}`}><div className="match-score"><div><strong>{match.Casa}</strong></div><div className="match-score-result"><b>{match.Risultato || '—'}</b><span>{match.PuntiCasa} - {match.PuntiTrasferta}</span></div><div><strong>{match.Trasferta}</strong></div></div><div className="match-teams"><PlayerMatchList team={match.Casa} side="home" rows={dayData.players[match.Casa] || []} bonuses={{ Difesa: match.ModificatoreDifesaCasa, Centrocampo: match.ModificatoreCentrocampoCasa, Attacco: match.ModificatoreAttaccoCasa }} /><PlayerMatchList team={match.Trasferta} side="away" rows={dayData.players[match.Trasferta] || []} bonuses={{ Difesa: match.ModificatoreDifesaTrasferta, Centrocampo: match.ModificatoreCentrocampoTrasferta, Attacco: match.ModificatoreAttaccoTrasferta }} /></div></article>)}</div>)}</div>
}

function UserStats({ history, users }) {
  const winningUsers = new Map()
  const addWinner = (user) => winningUsers.set(user, (winningUsers.get(user) || 0) + 1)
  const ownersOf = (row) => row?.Utenti?.length ? row.Utenti : row?.Utente ? [row.Utente] : []
  history.forEach((seasonData) => {
    for (const rows of [seasonData.championship, seasonData.fantasy, seasonData.scorers]) {
      ownersOf(rows?.[0]).forEach(addWinner)
    }
    for (const cup of seasonData.cups || []) {
      cup.positions.filter((entry) => entry.position === 1).flatMap((entry) => entry.users).forEach(addWinner)
    }
  })
  const mostWinningUser = users.slice().sort((a, b) => (winningUsers.get(b) || 0) - (winningUsers.get(a) || 0) || a.localeCompare(b, 'it'))[0] || ''
  const [selected, setSelected] = useState(mostWinningUser)
  useEffect(() => setSelected(mostWinningUser), [mostWinningUser])
  const records = history.map((seasonData) => {
    const find = (rows) => rows.find((row) => row.Utenti?.includes(selected) || row.Utente === selected)
    const fantasy = find(seasonData.fantasy)
    const championship = find(seasonData.championship)
    const champions = find(seasonData.champions)
    const scorer = find(seasonData.scorers)
    const cups = (seasonData.cups || []).map((cup) => ({ ...cup, result: cup.positions.find((entry) => entry.users.includes(selected)) })).filter((cup) => cup.result)
    return { season: seasonData.season, fantasy, championship, champions, scorer, cups }
  }).filter(({ fantasy, championship, scorer, cups }) => fantasy || championship || scorer || cups.length)
  const rankOf = (rows, target) => {
    const rankedRows = target?.Girone ? rows.filter((row) => row.Girone === target.Girone) : rows
    const index = rankedRows.findIndex((row) => row.Utenti?.includes(selected) || row.Utente === selected)
    return index < 0 ? '—' : index + 1
  }
  const wins = records.flatMap(({ season, fantasy, championship, scorer, cups }) => [
    championship && rankOf(history.find((item) => item.season === season).championship, championship) === 1 ? `Campionato ${season}` : '',
    fantasy && rankOf(history.find((item) => item.season === season).fantasy, fantasy) === 1 ? `Fantapunti ${season}` : '',
    scorer && rankOf(history.find((item) => item.season === season).scorers, scorer) === 1 ? `Capocannonieri ${season}` : '',
    ...cups.filter((cup) => cup.result.position === 1).map((cup) => `${cup.label} ${season}`),
  ]).filter(Boolean)
  const cupResult = (cups, slug) => cups.find((cup) => cup.slug === slug)?.result
  return <div className="user-page"><div className="panel user-selector"><div className="panel-toolbar"><div><h2>Statistiche utente</h2><p>Confronta i risultati di tutte le stagioni</p></div></div><label>Seleziona utente<select value={selected} onChange={(event) => setSelected(event.target.value)}>{users.map((user) => <option key={user} value={user}>{user}</option>)}</select></label></div><div className="stats-grid"><Stat value={records.length} label="stagioni" /><Stat value={wins.length} label="vittorie" /><Stat value={records.filter(({ championship }) => championship).length} label="campionati giocati" /></div><div className="panel"><div className="panel-toolbar"><div><h2>Riepilogo stagioni</h2><p>Posizione nelle classifiche e nelle coppe</p></div></div><div className="table-wrap"><table className="user-history-table"><thead><tr><th>Stagione</th><th>Squadra</th><th>Campionato</th><th>Fantapunti</th><th>Capocannonieri</th><th>Champions</th><th>Europa League</th><th>Coppa Italia</th></tr></thead><tbody>{records.map(({ season, fantasy, championship, scorer, cups }) => { const seasonData = history.find((item) => item.season === season); const championshipRank = championship && rankOf(seasonData.championship, championship); const fantasyRank = fantasy && rankOf(seasonData.fantasy, fantasy); const scorerRank = scorer && rankOf(seasonData.scorers, scorer); return <tr key={season}><td className="season-cell">{season.replace('_', '/')}</td><td className="team-name">{fantasy?.Squadra || championship?.Squadra || scorer?.Squadra || cups[0]?.result.team}</td><td>{championship ? <HistoryResult rank={championshipRank} value={`${championship.Punti} pt`} /> : '—'}</td><td>{fantasy ? <HistoryResult rank={fantasyRank} value={fantasy.FantaPunti} /> : '—'}</td><td>{scorer ? <HistoryResult rank={scorerRank} value={scorer.FantaPunti} /> : '—'}</td><td><CupResult result={cupResult(cups, 'champions')} /></td><td><CupResult result={cupResult(cups, 'europa_league')} /></td><td><CupResult result={cupResult(cups, 'coppa_italia')} /></td></tr> })}</tbody></table></div></div><div className="panel wins-panel"><div className="panel-toolbar"><div><h2>Vittorie</h2><p>Prime posizioni nelle classifiche disponibili</p></div></div>{wins.length ? <div className="win-list">{wins.map((win) => <span key={win}>{win}</span>)}</div> : <p className="empty">Nessuna vittoria registrata.</p>}</div></div>
}

function Stat({ value, label }) { return <div className="stat-card"><strong>{value}</strong><span>{label}</span></div> }

function Roster({ rows, teams, users }) {
  const [selected, setSelected] = useState(teams[0] || '')
  useEffect(() => setSelected(teams[0] || ''), [teams])
  const roster = rows.filter((row) => row.squadra === selected)
  const owner = [...new Set(users.filter((row) => normalizeTeam(row.Squadra) === normalizeTeam(selected)).map((row) => String(row.Utente || '').trim()).filter(Boolean))].join(' / ') || '—'
  const maxQuotation = Math.max(...roster.map((player) => numeric(player.quotazione)), 0)
  const clubCounts = Object.entries(roster.reduce((counts, player) => {
    const club = String(player.club || '').trim() || 'Squadra non indicata'
    counts[club] = (counts[club] || 0) + 1
    return counts
  }, {})).sort(([, countA], [, countB]) => countB - countA)
  return <div className="roster-grid"><div className="panel roster-selector"><div className="panel-toolbar"><div><h2>Seleziona una rosa</h2><p>Consulta i giocatori di ogni squadra</p></div></div><label>Squadra<select value={selected} onChange={(event) => setSelected(event.target.value)}>{teams.map((team) => <option key={team} value={team}>{team}</option>)}</select></label><div className="roster-owner"><span>Utente</span><strong>{owner}</strong></div><div className="roster-summary"><span className="big-number">{roster.length}</span><span>giocatori<br />in rosa</span></div>{clubCounts.length > 0 && <div className="club-counts"><p>Giocatori per club</p>{clubCounts.map(([club, count]) => <div className="club-count" key={club}><span>{club}</span><strong>{count}</strong></div>)}</div>}</div><div className="panel roster-list"><div className="panel-toolbar"><div><h2>{selected || 'Rosa'}</h2><p>Utente: {owner} · Rosa completa</p></div></div>{roster.length ? <div className="players">{roster.map((player, index) => { const role = String(player.ruolo).trim().toLowerCase().charAt(0); const roleClass = { p: 'por', d: 'dif', c: 'cen', a: 'att' }[role] || role; return <div className={`player player-${roleClass}`} key={`${player.giocatore}-${index}`}><span className={`role role-${roleClass}`}>{player.ruolo || '—'}</span><span className="player-details"><span>{player.giocatore}</span>{player.club && <small>{player.club}</small>}</span>{player.quotazione && <strong className={numeric(player.quotazione) === maxQuotation ? 'best-value' : ''}>{player.quotazione}</strong>}</div> })}</div> : <p className="empty">Nessun giocatore trovato per questa squadra.</p>}</div></div>
}

createRoot(document.getElementById('root')).render(<App />)
