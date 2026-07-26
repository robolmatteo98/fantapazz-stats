import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import './styles.css'

const dataFiles = import.meta.glob('../data/*/*', { eager: true, as: 'url' })
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

const navItems = [
  { id: 'fantasy', label: 'Fantapunti', icon: '✦' },
  { id: 'championship', label: 'Campionato', icon: '⌁' },
  { id: 'scorers', label: 'Capocannonieri', icon: '⚽' },
  { id: 'rosters', label: 'Rose', icon: '♟' },
]

function parseCsv(text) {
  const lines = text.replace(/^\uFEFF/, '').trim().split(/\r?\n/)
  const split = (line) => line.match(/(?:"([^"]*)")|([^;]+)/g)?.map((cell) => cell.replace(/^"|"$/g, '').trim()) || []
  const headers = split(lines[0])
  return lines.slice(1).filter(Boolean).map((line) => Object.fromEntries(split(line).map((value, i) => [headers[i], value])))
}

function numeric(value) {
  return Number(String(value ?? '').replace(',', '.')) || 0
}

function App() {
  const [page, setPage] = useState('fantasy')
  const [season, setSeason] = useState(availableSeasons[availableSeasons.length - 1] || '')
  const seasons = availableSeasons
  const [data, setData] = useState({ fantasy: [], championship: [], scorers: [], rosters: [] })
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
        const rosterBuffer = contents.rose
        if (!fantasyText || !scorersText || !rosterBuffer) throw new Error(`File mancanti nella stagione ${season}`)
        // SheetJS is more reliable in the browser when given a typed byte array.
        const workbook = XLSX.read(new Uint8Array(rosterBuffer), { type: 'array' })
        const rosters = workbook.SheetNames.flatMap((sheetName) => {
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
        })
        setData({ fantasy: parseCsv(fantasyText), championship: championshipText ? parseCsv(championshipText) : [], scorers: parseCsv(scorersText), rosters })
      } catch (err) {
        setError(err.message)
      } finally { setLoading(false) }
    }
    load()
  }, [season])

  useEffect(() => {
    const currentNav = navItems.find((item) => item.id === page)
    if (currentNav?.id === 'championship' && !data.championship.length) setPage('fantasy')
  }, [data.championship, page])

  const teams = useMemo(() => [...new Set(data.rosters.map((row) => row.squadra).filter(Boolean))], [data.rosters])
  const current = navItems.find((item) => item.id === page)
  const visibleNavItems = navItems.filter((item) => item.id !== 'championship' || data.championship.length)
  const seasonLabel = season.replace('_', '/')

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><span className="brand-mark">F</span><span>Fantapazz<br /><small>STATS</small></span></div>
      <label className="season"><span className="live-dot" /> Stagione<select value={season} onChange={(event) => setSeason(event.target.value)}>{seasons.map((value) => <option key={value} value={value}>{value.replace('_', '/')}</option>)}</select></label>
      <nav>{visibleNavItems.map((item) => <button key={item.id} className={page === item.id ? 'active' : ''} onClick={() => setPage(item.id)}><span>{item.icon}</span>{item.label}</button>)}</nav>
      <div className="sidebar-foot">DATI LOCALI<br /><span>Importati dai file CSV</span></div>
    </aside>
    <main className="content">
      <header className="topbar"><div className="mobile-brand"><span className="brand-mark">F</span> Fantapazz Stats</div><div className="top-season">{seasonLabel} <span>•</span> Lega Fantapazz</div><label className="mobile-season"><span>Stagione</span><select aria-label="Seleziona stagione" value={season} onChange={(event) => setSeason(event.target.value)}>{seasons.map((value) => <option key={value} value={value}>{value.replace('_', '/')}</option>)}</select></label><div className="avatar">FP</div></header>
      <section className="page-heading"><div><p className="eyebrow">PANORAMICA DELLA LEGA</p><h1>{current.label}</h1><p className="subheading">Classifica e statistiche della stagione {seasonLabel}</p></div><div className="data-badge"><span className="pulse" /> Dati aggiornati</div></section>
      {loading && <div className="state-card">Caricamento dati…</div>}
      {error && <div className="state-card error">{error}. Avvia l’app tramite <code>npm run dev</code> per servire i file locali.</div>}
      {!loading && !error && page === 'fantasy' && <Ranking rows={data.fantasy} type="fantasy" />}
      {!loading && !error && page === 'championship' && <Ranking rows={data.championship} type="championship" />}
      {!loading && !error && page === 'scorers' && <Ranking rows={data.scorers} type="scorers" />}
      {!loading && !error && page === 'rosters' && <Roster rows={data.rosters} teams={teams} />}
    </main>
  </div>
}

function Ranking({ rows, type }) {
  const isChampionship = type === 'championship'
  const isScorers = type === 'scorers'
  const scoreLabel = isChampionship ? 'Punti' : isScorers ? 'Punti gol' : 'Fantapunti'
  const scoreKey = isChampionship ? 'Punti' : 'FantaPunti'
  const title = isChampionship ? 'Classifica campionato' : isScorers ? 'Classifica capocannonieri' : 'Classifica fantapunti'
  return <div className="panel"><div className="panel-toolbar"><div><h2>{title}</h2><p>{rows.length} squadre partecipanti</p></div><div className="count-pill">{rows.length} <span>squadre</span></div></div><div className="table-wrap"><table><thead><tr><th>#</th><th>Squadra</th><th>PG</th><th>V</th><th>N</th><th>P</th><th>GF</th><th>GS</th><th className="right">{scoreLabel}</th></tr></thead><tbody>{rows.map((row, index) => <tr key={`${row.Squadra}-${index}`}><td><Rank rank={index + 1} /></td><td className="team-name">{row.Squadra}</td><td>{row.Punti}</td><td>{row.Vittorie}</td><td>{row.Nulle}</td><td>{row.Sconfitte}</td><td>{row.GolFa}</td><td>{row.GolSu}</td><td className="score">{numeric(row[scoreKey]).toLocaleString('it-IT', { minimumFractionDigits: String(row[scoreKey]).includes(',') ? 1 : 0 })}</td></tr>)}</tbody></table></div></div>
}

function Rank({ rank }) { return rank <= 3 ? <span className={`medal medal-${rank}`}>{rank}</span> : <span className="rank">{rank}</span> }

function Roster({ rows, teams }) {
  const [selected, setSelected] = useState(teams[0] || '')
  useEffect(() => setSelected(teams[0] || ''), [teams])
  const roster = rows.filter((row) => row.squadra === selected)
  const clubCounts = Object.entries(roster.reduce((counts, player) => {
    const club = String(player.club || '').trim() || 'Squadra non indicata'
    counts[club] = (counts[club] || 0) + 1
    return counts
  }, {})).sort(([, countA], [, countB]) => countB - countA)
  return <div className="roster-grid"><div className="panel roster-selector"><div className="panel-toolbar"><div><h2>Seleziona una rosa</h2><p>Consulta i giocatori di ogni squadra</p></div></div><label>Squadra<select value={selected} onChange={(event) => setSelected(event.target.value)}>{teams.map((team) => <option key={team} value={team}>{team}</option>)}</select></label><div className="roster-summary"><span className="big-number">{roster.length}</span><span>giocatori<br />in rosa</span></div>{clubCounts.length > 0 && <div className="club-counts"><p>Giocatori per club</p>{clubCounts.map(([club, count]) => <div className="club-count" key={club}><span>{club}</span><strong>{count}</strong></div>)}</div>}</div><div className="panel roster-list"><div className="panel-toolbar"><div><h2>{selected || 'Rosa'}</h2><p>Rosa completa</p></div></div>{roster.length ? <div className="players">{roster.map((player, index) => { const role = String(player.ruolo).toLowerCase().slice(0, 3); return <div className={`player player-${role}`} key={`${player.giocatore}-${index}`}><span className={`role role-${role}`}>{player.ruolo || '—'}</span><span className="player-details"><span>{player.giocatore}</span>{player.club && <small>{player.club}</small>}</span>{player.quotazione && <strong>{player.quotazione}</strong>}</div> })}</div> : <p className="empty">Nessun giocatore trovato per questa squadra.</p>}</div></div>
}

createRoot(document.getElementById('root')).render(<App />)
