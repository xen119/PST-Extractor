import '@testing-library/jest-dom/vitest'
import { afterEach, vi } from 'vitest'
import { cleanup } from '@testing-library/react'

class ResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
}

class MemoryStorage implements Storage {
  private readonly store = new Map<string, string>()

  get length(): number {
    return this.store.size
  }

  clear(): void {
    this.store.clear()
  }

  getItem(key: string): string | null {
    return this.store.has(key) ? this.store.get(key) || null : null
  }

  key(index: number): string | null {
    return Array.from(this.store.keys())[index] || null
  }

  removeItem(key: string): void {
    this.store.delete(key)
  }

  setItem(key: string, value: string): void {
    this.store.set(key, String(value))
  }
}

Object.defineProperty(window, 'ResizeObserver', {
  value: ResizeObserver
})

Object.defineProperty(window, 'matchMedia', {
  value: vi.fn().mockImplementation(() => ({
    matches: false,
    media: '',
    onchange: null,
    addEventListener: vi.fn(),
    removeEventListener: vi.fn(),
    addListener: vi.fn(),
    removeListener: vi.fn(),
    dispatchEvent: vi.fn()
  }))
})

Object.defineProperty(window, 'localStorage', {
  value: new MemoryStorage()
})

Object.defineProperty(window, 'sessionStorage', {
  value: new MemoryStorage()
})

window.scrollTo = vi.fn()
window.HTMLElement.prototype.scrollIntoView = vi.fn()

afterEach(() => {
  cleanup()
  document.body.style.pointerEvents = ''
  document.body.style.overflow = ''
  document.body.removeAttribute('data-scroll-locked')
})
