#!/usr/bin/env node

process.argv.push('--competizione', 'campionato')
await import('./download_giornate.js')
