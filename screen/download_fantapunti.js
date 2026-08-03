#!/usr/bin/env node

process.argv.push('--competizione', 'fantapunti', '--solo-risultati')
await import('./download_giornate.js')
