#!/usr/bin/env node

/*
 * Scarica le giornate della competizione Campionato di Fantapazz.
 *
 * Esempio:
 *   node screen/scarica-giornate.js \
 *     --anno 18_19 \
 *     --url https://www.fantapazz.com/fantacalcio/formazioni-in-campo-lega/460/1/113304
 *
 * Se la lega richiede autenticazione, passa il cookie senza inserirlo nel file:
 *   FANTAPAZZ_COOKIE='...cookie...' node screen/scarica-giornate.js --anno 18_19 --url '...'
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
const startUrl = argument('--url')
if (!year || !startUrl) {
  console.error('Uso: node screen/scarica-giornate.js --anno 18_19 --url "URL_GIORNATA_1"')
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

function extractCampionatoUrl(optionValue) {
  try {
    const url = new URL(optionValue, startUrl)
    const parts = url.pathname.split('/').filter(Boolean)
    // /fantacalcio/formazioni-in-campo-lega/{giornataId}/1/{legaId}
    return parts.at(-2) === '1' ? url.href : ''
  } catch {
    return ''
  }
}

function extractDays(indexHtml) {
  const days = []
  const optionPattern = /<option\b[^>]*value=["']([^"']+)["'][^>]*>([\s\S]*?)<\/option>/gi
  for (const match of indexHtml.matchAll(optionPattern)) {
    const url = extractCampionatoUrl(match[1])
    if (!url || days.some((day) => day.url === url)) continue
    const label = textFromHtml(match[2])
    const dayMatch = label.match(/Giornata\s+(\d+)/i)
    if (dayMatch) days.push({ number: Number(dayMatch[1]), url })
  }
  return days.sort((a, b) => a.number - b.number)
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

function safeFolderName(value) {
  return String(value || 'squadra').trim().replace(/[<>:"/\\|?*\u0000-\u001F]/g, '-').replace(/\s+/g, ' ')
}

function extractPlayers(html, dayNumber, date) {
  const formationBlocks = [...html.matchAll(/<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b[^"']*["'][^>]*>([\s\S]*?)(?=<div\b[^>]*class=["'][^"']*\bformazioni1vs1\b|$)/gi)]
  const players = []
  for (const formationMatch of formationBlocks) {
    const formation = formationMatch[1]
    const teams = [...formation.matchAll(/<div\b[^>]*class=["'][^"']*\balias\b[^"']*["'][^>]*>([\s\S]*?)<\/div>/gi)].map((match) => textFromHtml(match[1]))
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
        Squadra: teams[side === 'away' ? 1 : 0] || '',
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
  console.log(`Leggo le giornate da ${startUrl}`)
  const indexHtml = await getHtml(startUrl)
  const days = extractDays(indexHtml)
  if (!days.length) throw new Error('Nessuna giornata di Campionato trovata nel menu della pagina')

  const outputDir = join('data', year, 'campionato')
  await mkdir(outputDir, { recursive: true })
  console.log(`Trovate ${days.length} giornate`)

  for (const [index, day] of days.entries()) {
    const html = index === 0 ? indexHtml : await getHtml(day.url)
    const matches = extractMatches(html, day.number)
    const date = textFromHtml(html.match(/<div class=["']orario["'][^>]*>([\s\S]*?)<\/div>/i)?.[1] || '')
    const players = extractPlayers(html, day.number, date)
    if (!matches.length) {
      console.warn(`Giornata ${day.number}: nessuna partita trovata`)
      continue
    }
    const headers = ['Giornata', 'Data', 'Casa', 'PuntiCasa', 'Risultato', 'Trasferta', 'PuntiTrasferta']
    const csv = [headers, ...matches.map((match) => headers.map((header) => match[header]))]
      .map((row) => row.map(csvCell).join(';'))
      .join('\n') + '\n'
    const dayDir = join(outputDir, `${day.number}_giornata`)
    await mkdir(dayDir, { recursive: true })
    const outputPath = join(dayDir, 'totale.csv')
    await writeFile(outputPath, `\uFEFF${csv}`, 'utf8')
    console.log(`Giornata ${day.number}: ${matches.length} partite → ${outputPath}`)

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
    console.log(`  Giocatori estratti: ${players.length} in ${playersByTeam.size} squadre`)
  }
}

main().catch((error) => {
  console.error(`Errore: ${error.message}`)
  console.error('Se la pagina richiede login, usa FANTAPAZZ_COOKIE con i cookie della sessione del browser.')
  process.exit(1)
})
