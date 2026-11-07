import 'dotenv/config'
import express from 'express'
import multer from 'multer'
import { spawn } from 'node:child_process'
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import path from 'node:path'
import crypto from 'node:crypto'

const app = express()
const port = Number(process.env.PORT || 3000)
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } })
app.use(express.json({ limit: '1mb' }))

function runQpdf(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.env.QPDF_PATH || 'qpdf', args, { windowsHide: true })
    let stderr = ''
    child.stderr.on('data', d => { stderr += d.toString() })
    child.on('error', err => reject(new Error(`Could not start qpdf. Install qpdf and ensure it is on PATH. ${err.message}`)))
    child.on('close', code => code === 0 ? resolve() : reject(new Error(stderr.trim() || `qpdf exited with code ${code}`)))
  })
}

app.post('/api/pdf/password', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Choose a PDF file.' })
  const { action, password } = req.body || {}
  if (!['encrypt', 'decrypt'].includes(action)) return res.status(400).json({ error: 'Invalid password action.' })
  if (typeof password !== 'string' || password.length < 1 || password.length > 256) return res.status(400).json({ error: 'Enter a valid password (1–256 characters).' })
  const dir = await mkdtemp(path.join(tmpdir(), 'ourpdf-'))
  const input = path.join(dir, `${crypto.randomUUID()}.pdf`)
  const output = path.join(dir, `${crypto.randomUUID()}.pdf`)
  try {
    await writeFile(input, req.file.buffer, { mode: 0o600 })
    if (action === 'encrypt') {
      // AES-256 encryption. The same supplied password is used for opening and owner permissions.
      await runQpdf(['--encrypt', password, password, '256', '--', input, output])
    } else {
      await runQpdf([`--password=${password}`, '--decrypt', input, output])
    }
    const result = await readFile(output)
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="${action === 'encrypt' ? 'ourpdf-protected' : 'ourpdf-unlocked'}.pdf"`)
    res.send(result)
  } catch (error) {
    const message = String(error?.message || error)
    res.status(400).json({ error: message.includes('invalid password') || message.includes('password') ? 'The password may be incorrect, or the PDF uses unsupported encryption.' : message })
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => {})
  }
})

app.post('/api/ai', async (req, res) => {
  const apiKey = process.env.GEMINI_API_KEY
  if (!apiKey) return res.status(503).json({ error: 'Gemini is not configured. Add GEMINI_API_KEY to the server .env file and restart the server.' })
  const { action, prompt, text } = req.body || {}
  if (!['summarize', 'chat', 'translate'].includes(action)) return res.status(400).json({ error: 'Invalid AI action.' })
  if (typeof prompt !== 'string' || typeof text !== 'string' || !text.trim()) return res.status(400).json({ error: 'A PDF and a prompt are required.' })
  if (text.length > 90000) return res.status(413).json({ error: 'This PDF contains too much text for one request. Try a shorter document.' })
  try {
    const model = process.env.GEMINI_MODEL || 'gemini-2.5-flash'
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: 'You are OurPDF assistant. Answer accurately using the supplied document only. If the answer is not in the document, say so. Do not invent citations or facts.' }] },
        contents: [{ role: 'user', parts: [{ text: `${prompt}\n\nDOCUMENT TEXT:\n${text}` }] }],
        generationConfig: { temperature: 0.25, maxOutputTokens: 4096 }
      })
    })
    const data = await response.json()
    if (!response.ok) {
      const detail = data?.error?.message || `Gemini returned HTTP ${response.status}`
      return res.status(response.status === 429 ? 429 : 502).json({ error: detail })
    }
    const answer = data?.candidates?.[0]?.content?.parts?.map(p => p.text || '').join('').trim()
    if (!answer) return res.status(502).json({ error: 'Gemini did not return a text response. The document may have triggered a safety or output limit.' })
    res.json({ text: answer })
  } catch (error) {
    res.status(502).json({ error: `Could not reach Gemini: ${error.message}` })
  }
})

// In production, serve the Vite build from this same origin.
app.use(express.static(path.resolve('dist')))
app.get('*', (req, res, next) => {
  if (req.path.startsWith('/api/')) return next()
  res.sendFile(path.resolve('dist/index.html'))
})
app.listen(port, () => console.log(`OurPDF server running at http://localhost:${port}`))
