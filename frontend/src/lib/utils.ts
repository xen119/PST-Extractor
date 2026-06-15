import { clsx, type ClassValue } from 'clsx'
import { twMerge } from 'tailwind-merge'

export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs))
}

export function normalizeText(value: unknown): string {
  return String(value ?? '').trim().replace(/\s+/g, ' ')
}

export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return 'Unknown size'
  }

  if (bytes < 1024) {
    return `${bytes} B`
  }

  const units = ['KB', 'MB', 'GB', 'TB']
  let value = bytes / 1024
  let unit = units[0]
  for (const nextUnit of units) {
    unit = nextUnit
    if (value < 1024) {
      break
    }
    value /= 1024
  }

  return `${value.toFixed(value >= 10 ? 0 : 1)} ${unit}`
}

export function formatDate(value: string | number | Date | null | undefined): string {
  if (value == null || value === '') {
    return 'Unknown'
  }

  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) {
    return String(value)
  }

  return new Intl.DateTimeFormat(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short'
  }).format(date)
}

export function getInitials(name: string, fallback = '?'): string {
  const text = normalizeText(name)
  if (!text) {
    return fallback
  }

  const parts = text.split(/\s+/).filter(Boolean)
  if (!parts.length) {
    return fallback
  }

  return parts
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() || '')
    .join('') || fallback
}

export function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => {
    switch (char) {
      case '&':
        return '&amp;'
      case '<':
        return '&lt;'
      case '>':
        return '&gt;'
      case '"':
        return '&quot;'
      case '\'':
        return '&#39;'
      default:
        return char
    }
  })
}

export function buildHtmlFrameSrcDoc(html: string, isDark: boolean): string {
  const frameText = isDark ? '#e6edf7' : '#1f2a37'
  const frameBackground = isDark ? '#111a29' : '#ffffff'
  const frameLink = isDark ? '#7ab8ff' : '#2b6cb0'
  const frameQuoteBorder = isDark ? '#3a4760' : '#c2ccd9'
  const frameQuoteText = isDark ? '#a8b3c4' : '#526072'

  return `<!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <base target="_blank" />
        <style>
          :root { color-scheme: ${isDark ? 'dark' : 'light'}; }
          body {
            margin: 0;
            padding: 8px 4px 12px;
            font-family: 'Segoe UI', Tahoma, Arial, sans-serif;
            font-size: 13px;
            line-height: 1.55;
            color: ${frameText};
            background: ${frameBackground};
          }
          img { max-width: 100%; height: auto; }
          table { border-collapse: collapse; }
          a { color: ${frameLink}; }
          blockquote {
            border-left: 3px solid ${frameQuoteBorder};
            margin-left: 0;
            padding-left: 1rem;
            color: ${frameQuoteText};
          }
        </style>
      </head>
      <body>${html}</body>
    </html>`
}

export function downloadTextFile(fileName: string, contents: string): void {
  const blob = new Blob([contents], { type: 'text/plain;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.style.display = 'none'
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}
