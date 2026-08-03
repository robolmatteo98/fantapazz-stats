#!/usr/bin/env node

process.argv.push('--competizione', 'coppa_italia')
await import('./download_giornate.js')
