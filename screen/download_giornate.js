#!/usr/bin/env node

/*
 * Scarica le giornate del Campionato e dei Fantapunti di Fantapazz.
 *
 * Esempio:
 *   node screen/download_giornate.js \
 *     --anno 18_19 \
 *     --url https://www.fantapazz.com/fantacalcio/formazioni-in-campo-lega/460/1/113304
 *
 * Se la lega richiede autenticazione, passa il cookie senza inserirlo nel file:
 *   FANTAPAZZ_COOKIE='...cookie...' node screen/download_giornate.js --anno 18_19 --url '...'
 * Oppure crea screen/.env (ignorato da Git) con:
 *   FANTAPAZZ_COOKIE=...
 */

import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import https from 'node:https'

const args = process.argv.slice(2)
const argument = (name) => {
  const index = args.indexOf(name)
  return index >= 0 ? args[index + 1] : ''
}

const year = argument('--anno')
const requestedCompetition = argument('--competizione')
const stage = argument('--fase')
const onlyResults = args.includes('--solo-risultati')
const competitionCodes = { campionato: 1, fantapunti: 6, champions: 4, europa_league: 5, coppa_italia: 20 }
const jornadaId = argument('--giornata')
const leagueId = argument('--lega')
let startUrl = argument('--url')
if (!startUrl && requestedCompetition && jornadaId && leagueId && competitionCodes[requestedCompetition]) {
  startUrl = `https://www.fantapazz.com/fantacalcio/formazioni-in-campo-lega/${jornadaId}/${competitionCodes[requestedCompetition]}/${leagueId}`
}
if (!year || !startUrl) {
  console.error('Uso: node screen/download_giornate.js --anno 18_19 --url "URL_GIORNATA_1" oppure --giornata 566 --lega 201910')
  process.exit(1)
}

let baseHeaders = {}

async function readEnvCookie() {
  try {
    const envPath = join(fileURLToPath(new URL('.', import.meta.url)), '.env')
    const content = await readFile(envPath, 'utf8')
    const value = content.match(/^\s*FANTAPAZZ_COOKIE\s*=\s*(.*?)\s*$/m)?.[1] || ''
    return value.replace(/^(["'])(.*)\1$/, '$2')
  } catch {
    return ''
  }
}

function decodeHtml(value) {
  return value
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([\da-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
}

function textFromHtml(value) {
  return decodeHtml(value.replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim()
}

function csvCell(value) {
  const text = String(value ?? '')
  return /[;"\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

function getHtml(url, redirects = 0) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, { headers: baseHeaders }, (response) => {
      const status = response.statusCode || 0
      if (status >= 300 && status < 400 && response.headers.location && redirects < 5) {
        response.resume()
        const nextUrl = new URL(response.headers.location, url).href
        resolve(getHtml(nextUrl, redirects + 1))
        return
      }
      if (status < 200 || status >= 300) {
        response.resume()
        reject(new Error(`${status} — ${url}`))
        return
      }
      let body = ''
      response.setEncoding('utf8')
      response.on('data', (chunk) => { body += chunk })
      response.on('end', () => resolve(body))
    })
    request.on('error', reject)
  })
}

function extractCompetitionUrl(optionValue, competitionId) {
  try {
    const url = new URL(optionValue, startUrl)
    const parts = url.pathname.split('/').filter(Boolean)
    // /fantacalcio/formazioni-in-campo-lega/{giornataId}/{competizione}/{legaId}
    return parts.at(-2) === String(competitionId) ? url.href : ''
  } catch {
    return ''
  }
}

function extractDays(indexHtml, competitionId) {
  const days = []
  const optionPattern = /<option\b[^>]*value=["']([^"']+)["'][^>]*>([\s\S]*?)<\/option>/gi
  for (const match of indexHtml.matchAll(optionPattern)) {
    const url = extractCompetitionUrl(match[1], competitionId)
    if (!url || days.some((day) => day.url === url)) continue
    const label = textFromHtml(match[2])
    const dayMatch = label.match(/Giornata\s+(\d+)/i)
    if (dayMatch) days.push({ number: Number(dayMatch[1]), url })
  }
  return days.sort((a, b) => a.number - b.number)
}

function competitionStartUrl(competitionId) {
  const url = new URL(startUrl)
  const parts = url.pathname.split('/')
  if (parts.length < 3) throw new Error(`URL non riconosciuto: ${startUrl}`)
  parts[parts.length - 2] = String(competitionId)
  url.pathname = parts.join('/')
  return url.href
}

function competitionIdFromUrl(urlValue) {
  const parts = new URL(urlValue).pathname.split('/').filter(Boolean)
  return Number(parts.at(-2))
}

function matchdayIdFromUrl(urlValue) {
  const parts = new URL(urlValue).pathname.split('/').filter(Boolean)
  return Number(parts.at(-3))
}

function safeSlug(value) {
  return String(value || 'competizione').trim().toLowerCase()
    .replace(/[^a-z0-9àèéìòù_-]+/gi, '_')
    .replace(/^_+|_+$/g, '') || 'competizione'
}

function extractMatches(html, dayNumber) {
  const date = textFromHtml(html.match(/<div class=["']orario["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
  const rows = [...html.matchAll(/<tr\b[^>]*class=["'][^"']*\bfPartita\b[^"']*["'][^>]*>([\s\S]*?)<\/tr>/gi)]
  return rows.map((rowMatch) => {
    const row = rowMatch[1]
    const teams = [...row.matchAll(/<div\b[^>]*class=["'][^"']*\bfSquadra\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi)].map((match) => textFromHtml(match[1]))
    const points = [...row.matchAll(/<div\b[^>]*class=["'][^"']*\bfPuntiInfPartita\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi)].map((match) => textFromHtml(match[1]))
    const result = textFromHtml(row.match(/<td\b[^>]*class=["'][^"']*\bGol\b[^"']*["'][^>]*>([\s\S]*?)<\/td>/i)?.[1] || '')
    if (teams.length < 2) return null
    return {
      Giornata: dayNumber,
      Data: date,
      Casa: teams[0],
      PuntiCasa: points[0] || '',
      Risultato: result,
      Trasferta: teams[1],
      PuntiTrasferta: points[1] || '',
    }
  }).filter(Boolean)
}

function extractTeamBonuses(html) {
  const bonuses = new Map()
  const blocks = [
    ...html.matchAll(/<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b[^"']*["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b|$)/gi),
    ...html.matchAll(/<div\b[^>]*class=["'][^"']*\bFormazioniFloating\b[^"']*["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bFormazioniFloating\b|$)/gi),
  ]
  for (const blockMatch of blocks) {
    const block = blockMatch[1]
    const teams = [...block.matchAll(/<(?:div|h1)\b[^>]*class=["'][^"']*\balias\b[^"']*["'][^>]*>([\s\S]*?)<\/(?:div|h1)>/gi)].map((match) => textFromHtml(match[1]))
    const lists = [...block.matchAll(/<ul\b[^>]*class=["'][^"']*\blistaBonus\b[^"']*["'][^>]*>([\s\S]*?)<\/ul>/gi)]
    lists.forEach((listMatch, index) => {
      const team = teams[index]
      if (!team) return
      const values = { Difesa: '', Centrocampo: '', Attacco: '' }
      const items = [...listMatch[1].matchAll(/<li\b[^>]*class=["'][^"']*\bBonusSquadra\b[^"']*["'][^>]*>([\s\S]*?)<\/li>/gi)]
      for (const item of items) {
        const label = textFromHtml(item[1].match(/<span\b[^>]*class=["'][^"']*\bdescr\b[^"']*["'][^>]*>([\s\S]*?)<\/span>/i)?.[1] || '')
        const value = textFromHtml(item[1].match(/<span\b[^>]*class=["'][^"']*\bvalore\b[^"']*["'][^>]*>([\s\S]*?)<\/span>/i)?.[1] || '')
        if (/difesa/i.test(label)) values.Difesa = value
        else if (/centrocampo|centro\s+campo/i.test(label)) values.Centrocampo = value
        else if (/attacco/i.test(label)) values.Attacco = value
      }
      bonuses.set(team, values)
    })
  }
  return bonuses
}

async function extractFantapuntiScores(html, dayNumber, leagueId, matchdayId) {
  const date = textFromHtml(html.match(/<div class=["'][^"']*orario[^"']*["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
  const endpoint = new URL(`/getHtmlFClassificaGiornata/${leagueId}/${matchdayId}?pageNumber=1&pageSize=100`, startUrl).href
  const response = await getHtml(endpoint)
  let fragments
  try {
    fragments = JSON.parse(response)
  } catch {
    throw new Error(`Risposta Fantapunti non valida per la giornata ${dayNumber}`)
  }
  const scores = []
  for (const fragment of Array.isArray(fragments) ? fragments : []) {
    const rows = [...String(fragment).matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)]
    for (const rowMatch of rows) {
      const row = rowMatch[1]
      const team = textFromHtml(row.match(/<div\b[^>]*class=["'][^"']*\bsquadra\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
      const cells = [...row.matchAll(/<td\b[^>]*>([\s\S]*?)<\/td>/gi)].map((match) => textFromHtml(match[1]))
      const teamIndex = cells.findIndex((cell) => cell === team)
      const point = teamIndex >= 0 ? cells.slice(teamIndex + 1).find((cell) => /^-?\d+(?:[.,]\d+)?$/.test(cell)) : ''
      if (team && !scores.some((score) => score.Squadra === team)) {
        scores.push({ Giornata: dayNumber, Data: date, Squadra: team, FantaPunti: point || '' })
      }
    }
  }
  return scores
}

function safeFolderName(value) {
  return String(value || 'squadra').trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').replace(/\s+/g, ' ')
}

function extractPlayers(html, dayNumber, date) {
  const formationBlocks = [
    ...html.matchAll(/<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b[^"']*["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b|$)/gi),
    ...html.matchAll(/<div\b[^>]*class=["'][^"']*\bFormazioniFloating\b[^"']*["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bFormazioniFloating\b|$)/gi),
  ]
  const players = []
  for (const formationMatch of formationBlocks) {
    const formation = formationMatch[1]
    const teams = [...formation.matchAll(/<div\b[^>]*class=["'][^"']*\balias\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi)].map((match) => textFromHtml(match[1]))
    const formationTeam = textFromHtml(formation.match(/<h1\b[^>]*class=["'][^"']*\balias\b[^"']*["'][^>]*>([\s\S]*?)<\/h1>/i)?.[1] || '')
    const playerBlocks = [...formation.matchAll(/<div\b[^>]*class=["']([^"']*\bfic-calciatore\b[^"']*)["'][^>]*nid=["']([^"']*)["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bfic-calciatore\b|$)/gi)]
    for (const playerMatch of playerBlocks) {
      const classes = playerMatch[1]
      const body = playerMatch[3]
      const side = /\baway\b/i.test(classes) ? 'away' : 'home'
      const role = textFromHtml(body.match(/<span\b[^>]*class=["'][^"']*\bRuolo\b[^"']*["'][^>]*>([\s\S]*?)<\/span>/i)?.[1] || '')
      const name = textFromHtml(body.match(/<div\b[^>]*class=["'][^"']*\bnomeCalciatore\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
      const vote = textFromHtml(body.match(/<div\b[^>]*class=["'][^"']*\bvotoUfficiale\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
      const bonuses = [...body.matchAll(/<span\b[^>]*class=["']([^"']*\bimgBonus[^"']*)["'][^>]*>/gi)].map((match) => match[1].trim()).join('|')
      if (!name) continue
      players.push({
        Giornata: dayNumber,
        Data: date,
        Squadra: teams[side === 'away' ? 1 : 0] || formationTeam,
        CasaTrasferta: side === 'away' ? 'Trasferta' : 'Casa',
        Stato: /\bTf_no\b/i.test(classes) ? 'Riserva' : 'Titolare',
        Giocatore: name,
        Ruolo: role,
        Voto: vote,
        BonusMalus: bonuses,
        CalciatoreId: playerMatch[2],
      })
    }
  }
  return players
}

async function main() {
  const cookie = process.env.FANTAPAZZ_COOKIE || await readEnvCookie()
  baseHeaders = {
    'user-agent': 'Mozilla/5.0 (compatible; FantapazzStats/1.0)',
    accept: 'text/html,application/xhtml+xml',
    ...(cookie ? { cookie } : {}),
  }
  const competitions = requestedCompetition
    ? [{
      id: competitionIdFromUrl(startUrl),
      slug: safeSlug(requestedCompetition),
      label: requestedCompetition,
      isHeadToHead: competitionIdFromUrl(startUrl) !== 6,
      url: startUrl,
    }]
    : [
      { id: 1, slug: 'campionato', label: 'Campionato', isHeadToHead: true },
      { id: 6, slug: 'fantapunti', label: 'Fantapunti', isHeadToHead: false },
    ]

  for (const competition of competitions) {
    const competitionUrl = competition.url || competitionStartUrl(competition.id)
    console.log(`\nLeggo ${competition.label} da ${competitionUrl}`)
    const indexHtml = await getHtml(competitionUrl)
    const days = extractDays(indexHtml, competition.id)
    if (!days.length) days.push({ number: 1, url: competitionUrl })
    if (!days.length) {
      console.warn(`Nessuna giornata di ${competition.label} trovata nel menu della pagina`)
      continue
    }

    const outputDir = join('data', year, competition.slug, stage ? safeSlug(stage) : '')
    const directPhase = requestedCompetition && !['campionato', 'fantapunti'].includes(competition.slug) && days.length === 1
    await mkdir(outputDir, { recursive: true })
    console.log(`Trovate ${days.length} giornate di ${competition.label}`)

    for (const [index, day] of days.entries()) {
      const html = index === 0 ? indexHtml : await getHtml(day.url)
      const date = textFromHtml(html.match(/<div class=["'][^"']*orario[^"']*["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
      const matches = competition.isHeadToHead
        ? extractMatches(html, day.number)
        : await extractFantapuntiScores(html, day.number, new URL(startUrl).pathname.split('/').filter(Boolean).at(-1), matchdayIdFromUrl(day.url))
      const teamBonuses = extractTeamBonuses(html)
      const bonusFor = (team) => teamBonuses.get(team) || { Difesa: '', Centrocampo: '', Attacco: '' }
      const matchesWithBonuses = competition.isHeadToHead
        ? matches.map((match) => ({
          ...match,
          ModificatoreDifesaCasa: bonusFor(match.Casa).Difesa,
          ModificatoreCentrocampoCasa: bonusFor(match.Casa).Centrocampo,
          ModificatoreAttaccoCasa: bonusFor(match.Casa).Attacco,
          ModificatoreDifesaTrasferta: bonusFor(match.Trasferta).Difesa,
          ModificatoreCentrocampoTrasferta: bonusFor(match.Trasferta).Centrocampo,
          ModificatoreAttaccoTrasferta: bonusFor(match.Trasferta).Attacco,
        }))
        : matches.map((match) => ({
          ...match,
          ModificatoreDifesa: bonusFor(match.Squadra).Difesa,
          ModificatoreCentrocampo: bonusFor(match.Squadra).Centrocampo,
          ModificatoreAttacco: bonusFor(match.Squadra).Attacco,
        }))
      const players = onlyResults || competition.slug === 'fantapunti' ? [] : extractPlayers(html, day.number, date)
      if (!matches.length) {
        console.warn(`${competition.label} giornata ${day.number}: nessun dato trovato`)
        continue
      }
      const headers = competition.isHeadToHead
        ? ['Giornata', 'Data', 'Casa', 'PuntiCasa', 'Risultato', 'Trasferta', 'PuntiTrasferta', 'ModificatoreDifesaCasa', 'ModificatoreCentrocampoCasa', 'ModificatoreAttaccoCasa', 'ModificatoreDifesaTrasferta', 'ModificatoreCentrocampoTrasferta', 'ModificatoreAttaccoTrasferta']
        : ['Giornata', 'Data', 'Squadra', 'FantaPunti', 'ModificatoreDifesa', 'ModificatoreCentrocampo', 'ModificatoreAttacco']
      const csv = [headers, ...matchesWithBonuses.map((match) => headers.map((header) => match[header]))]
        .map((row) => row.map(csvCell).join(';'))
        .join('\n') + '\n'
      const dayDir = directPhase ? outputDir : join(outputDir, `${day.number}_giornata`)
      await mkdir(dayDir, { recursive: true })
      const outputPath = join(dayDir, 'totale.csv')
      await writeFile(outputPath, `\uFEFF${csv}`, 'utf8')
      console.log(`${competition.label} giornata ${day.number}: ${matchesWithBonuses.length} righe → ${outputPath}`)

      const playerHeaders = ['Stato', 'Giocatore', 'Ruolo', 'Voto', 'BonusMalus', 'CalciatoreId']
      const playersByTeam = new Map()
      for (const player of players) {
        if (!playersByTeam.has(player.Squadra)) playersByTeam.set(player.Squadra, [])
        playersByTeam.get(player.Squadra).push(player)
      }
      for (const [team, teamPlayers] of playersByTeam) {
        const teamDir = join(dayDir, safeFolderName(team))
        await mkdir(teamDir, { recursive: true })
        const playersCsv = [playerHeaders, ...teamPlayers.map((player) => playerHeaders.map((header) => player[header]))]
          .map((row) => row.map(csvCell).join(';'))
          .join('\n') + '\n'
        await writeFile(join(teamDir, 'giocatori.csv'), `\uFEFF${playersCsv}`, 'utf8')
      }
      console.log(onlyResults || competition.slug === 'fantapunti'
        ? '  Solo risultati delle squadre: giocatori non scaricati'
        : `  Giocatori estratti: ${players.length} in ${playersByTeam.size} squadre`)
    }
  }
}

main().catch((error) => {
  console.error(`Errore: ${error.message}`)
  console.error('Se la pagina richiede login, usa FANTAPAZZ_COOKIE con i cookie della sessione del browser.')
  process.exit(1)
})
