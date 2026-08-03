#!/usr/bin/env node

process.argv.push('--competizione', 'europa_league')
await import('./download_giornate.js')
