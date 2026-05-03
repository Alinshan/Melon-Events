// Dev launcher script for Electron + Vite (Windows compatible)
const { execSync, spawn } = require('child_process')
const path = require('path')

// Step 1: Compile Electron TypeScript (use shell:true for .cmd files on Windows)
console.log('Compiling Electron main process...')
execSync('tsc -p tsconfig.electron.json', {
  stdio: 'inherit',
  cwd: __dirname,
  shell: true
})
console.log('Compile done. Starting Electron...')

// Step 2: Launch Electron
const electronExe = require('electron')
const proc = spawn(String(electronExe), ['.'], {
  stdio: 'inherit',
  cwd: __dirname,
  env: { ...process.env, NODE_ENV: 'development' }
})

proc.on('close', (code) => {
  process.exit(code || 0)
})
