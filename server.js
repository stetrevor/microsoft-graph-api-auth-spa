const express = require('express')
const morgan = require('morgan')
const path = require('path')
const argv = require('yargs')
  .usage('Usage: $0 -p [PORT]')
  .alias('p', 'port')
  .describe('port', '(Optional) Port Number - default is 3000')
  .strict().argv

const DEFAULT_PORT = 3000

const app = express()

let port = DEFAULT_PORT
if (argv.p) port = argv.p

app.use(morgan('dev'))

app.use(
  '/lib',
  express.static(path.join(__dirname, '../../lib/msal-browser/lib'))
)

app.use(express.static('app'))

app.get('*', function (req, res) {
  res.sendFile(path.join(__dirname, '/index.html'))
})

app.listen(port)
console.log(`Listening on port ${port}...`)
