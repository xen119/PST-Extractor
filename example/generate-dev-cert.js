#!/usr/bin/env node

const fs = require('fs')
const net = require('net')
const os = require('os')
const path = require('path')
const { spawnSync } = require('child_process')

function parseArgs(argv) {
  const result = {
    outDir: 'certs',
    dnsNames: ['localhost'],
    days: 825,
    password: ''
  }

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i]
    if (arg === '--out' && argv[i + 1]) {
      result.outDir = argv[++i]
      continue
    }
    if (arg === '--dns' && argv[i + 1]) {
      result.dnsNames = argv[++i]
        .split(/[,\n;]/g)
        .map((item) => item.trim())
        .filter(Boolean)
      continue
    }
    if (arg === '--days' && argv[i + 1]) {
      const parsed = Number.parseInt(argv[++i], 10)
      if (Number.isInteger(parsed) && parsed > 0) {
        result.days = parsed
      }
      continue
    }
    if (arg === '--password' && argv[i + 1]) {
      result.password = String(argv[++i])
      continue
    }
    if (arg === '--help' || arg === '-h') {
      printHelp()
      process.exit(0)
    }
  }

  if (!result.dnsNames.length) {
    result.dnsNames = ['localhost']
  }

  return result
}

function printHelp() {
  console.log(`Usage:
  node generate-dev-cert.js [--out certs] [--dns localhost] [--days 825] [--password secret]

Generates a local HTTPS certificate bundle for the example server.
The helper writes a .pfx file that can be used with HTTPS_PFX_FILE.
`)
}

function sanitizeFileBase(value) {
  const text = String(value || '').trim().toLowerCase()
  const cleaned = text.replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '')
  return cleaned || 'localhost'
}

function resolveOutputDir(value) {
  const outputDir = path.resolve(__dirname, String(value || '').trim() || 'certs')
  fs.mkdirSync(outputDir, { recursive: true })
  return outputDir
}

function shellQuote(value) {
  return `'${String(value).replace(/'/g, "''")}'`
}

function runCommand(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: 'utf8',
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
    ...options
  })

  if (result.error) {
    throw result.error
  }

  if (result.status !== 0) {
    const stderr = String(result.stderr || '').trim()
    const stdout = String(result.stdout || '').trim()
    const message = [stderr, stdout].filter(Boolean).join('\n')
    throw new Error(message || `${command} exited with code ${result.status}`)
  }

  return result.stdout
}

function commandExists(command, args) {
  const result = spawnSync(command, args, {
    encoding: 'utf8',
    windowsHide: true,
    stdio: 'ignore'
  })
  return !result.error && result.status === 0
}

function resolvePowerShellCommand() {
  const commands = process.platform === 'win32'
    ? ['powershell.exe', 'powershell', 'pwsh']
    : ['pwsh', 'powershell', 'powershell.exe']

  for (const command of commands) {
    if (commandExists(command, ['-NoProfile', '-Command', '$PSVersionTable.PSVersion.Major'])) {
      return command
    }
  }

  return ''
}

function buildOpenSslConfig(dnsNames) {
  const subjectAltNames = []
  let dnsIndex = 1
  let ipIndex = 1

  for (const name of dnsNames) {
    if (net.isIP(name)) {
      subjectAltNames.push(`IP.${ipIndex} = ${name}`)
      ipIndex += 1
      continue
    }
    subjectAltNames.push(`DNS.${dnsIndex} = ${name}`)
    dnsIndex += 1
  }

  if (!subjectAltNames.length) {
    subjectAltNames.push('DNS.1 = localhost')
  }

  const commonName = dnsNames[0] || 'localhost'
  return `
[ req ]
default_bits = 2048
prompt = no
default_md = sha256
distinguished_name = req_distinguished_name
x509_extensions = v3_req

[ req_distinguished_name ]
CN = ${commonName}

[ v3_req ]
basicConstraints = critical,CA:FALSE
keyUsage = critical,digitalSignature,keyEncipherment
extendedKeyUsage = serverAuth
subjectAltName = @alt_names

[ alt_names ]
${subjectAltNames.join('\n')}
`.trim() + '\n'
}

function generateWithOpenSsl(options) {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pst-extractor-cert-'))
  const baseName = sanitizeFileBase(options.dnsNames[0] || 'localhost')
  const keyPath = path.join(tempDir, `${baseName}.key`)
  const crtPath = path.join(tempDir, `${baseName}.crt`)
  const pfxPath = path.join(tempDir, `${baseName}.pfx`)
  const configPath = path.join(tempDir, 'openssl.cnf')

  fs.writeFileSync(configPath, buildOpenSslConfig(options.dnsNames))
  try {
    runCommand(
      'openssl',
      [
        'req',
        '-x509',
        '-newkey',
        'rsa:2048',
        '-nodes',
        '-days',
        String(options.days),
        '-keyout',
        keyPath,
        '-out',
        crtPath,
        '-config',
        configPath,
        '-extensions',
        'v3_req'
      ],
      { cwd: tempDir }
    )

    runCommand(
      'openssl',
      [
        'pkcs12',
        '-export',
        '-out',
        pfxPath,
        '-inkey',
        keyPath,
        '-in',
        crtPath,
        '-passout',
        `pass:${options.password || ''}`
      ],
      { cwd: tempDir }
    )

    const outputDir = resolveOutputDir(options.outDir)
    const finalBase = path.join(outputDir, baseName)
    fs.copyFileSync(keyPath, `${finalBase}.key`)
    fs.copyFileSync(crtPath, `${finalBase}.crt`)
    fs.copyFileSync(pfxPath, `${finalBase}.pfx`)
    return {
      baseName,
      outputDir,
      pfxPath: `${finalBase}.pfx`,
      keyPath: `${finalBase}.key`,
      crtPath: `${finalBase}.crt`
    }
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true })
  }
}

function generateWithPowerShell(options) {
  const outputDir = resolveOutputDir(options.outDir)
  const baseName = sanitizeFileBase(options.dnsNames[0] || 'localhost')
  const pfxPath = path.join(outputDir, `${baseName}.pfx`)
  const dnsArray = options.dnsNames.map((dns) => shellQuote(dns)).join(', ')
  const scriptLines = [
    '$ErrorActionPreference = \'Stop\'',
    `$dnsNames = @(${dnsArray})`,
    `$pfxPath = ${shellQuote(pfxPath)}`,
    `$days = [int]${Number(options.days)}`,
    `$passwordText = ${shellQuote(options.password || '')}`,
    '$securePassword = New-Object System.Security.SecureString',
    'foreach ($char in $passwordText.ToCharArray()) {',
    '  [void]$securePassword.AppendChar($char)',
    '}',
    '$securePassword.MakeReadOnly()',
    '$rsa = [System.Security.Cryptography.RSA]::Create(2048)',
    'try {',
    '  $subject = [System.Security.Cryptography.X509Certificates.X500DistinguishedName]::new("CN=$($dnsNames[0])")',
    '  $certRequest = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new($subject, $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)',
    '  $san = [System.Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()',
    '  foreach ($name in $dnsNames) {',
    '    $parsedIp = $null',
    '    if ([System.Net.IPAddress]::TryParse($name, [ref]$parsedIp)) {',
    '      $san.AddIpAddress($parsedIp)',
    '    } else {',
    '      $san.AddDnsName($name)',
    '    }',
    '  }',
    '  $certRequest.CertificateExtensions.Add($san.Build())',
    '  $certRequest.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $true))',
    '  $keyUsage = [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature -bor [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyEncipherment',
    '  $certRequest.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new($keyUsage, $true))',
    '  $notBefore = [DateTimeOffset]::UtcNow.AddDays(-1)',
    '  $notAfter = [DateTimeOffset]::UtcNow.AddDays($days)',
    '  $cert = $certRequest.CreateSelfSigned($notBefore, $notAfter)',
    '  try {',
    '    $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $securePassword)',
    '    [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)',
    '  } finally {',
    '    $cert.Dispose()',
    '  }',
    '} finally {',
    '  $rsa.Dispose()',
    '}'
  ]
  const script = `${scriptLines.join('\n')}\n`

  const command = resolvePowerShellCommand()
  if (!command) {
    throw new Error('PowerShell is not available to generate a certificate.')
  }

  runCommand(command, [
    '-NoProfile',
    '-ExecutionPolicy',
    'Bypass',
    '-Command',
    script
  ])

  return {
    baseName,
    outputDir,
    pfxPath
  }
}

function main() {
  const options = parseArgs(process.argv.slice(2))
  const hasOpenSsl = commandExists('openssl', ['version'])
  const generated = hasOpenSsl ? generateWithOpenSsl(options) : generateWithPowerShell(options)
  const port = Number.parseInt(process.env.PORT || '3030', 10) || 3030
  const relativeOutput = path.relative(__dirname, generated.outputDir) || '.'
  const pfxRelative = path.relative(__dirname, generated.pfxPath)

  console.log(`Generated HTTPS certificate bundle:
- ${path.join(relativeOutput, `${generated.baseName}.pfx`)}`)
  if (generated.keyPath && generated.crtPath) {
    console.log(`- ${path.join(relativeOutput, `${generated.baseName}.key`)}`)
    console.log(`- ${path.join(relativeOutput, `${generated.baseName}.crt`)}`)
  }
  console.log('')
  console.log('Use these settings in example/.env:')
  console.log('HTTPS_ENABLED=true')
  console.log(`HTTPS_PFX_FILE=${pfxRelative}`)
  console.log(`PUBLIC_BASE_URL=https://localhost:${port}`)
}

main()
