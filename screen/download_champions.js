#!/usr/bin/env node

process.argv.push('--competizione', 'champions')
await import('./download_giornate.js')
