import '@testing-library/jest-dom/vitest'
import { afterEach, vi } from 'vitest'
import { cleanup } from '@testing-library/react'

class ResizeObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
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

window.scrollTo = vi.fn()
window.HTMLElement.prototype.scrollIntoView = vi.fn()

afterEach(() => {
  cleanup()
  document.body.style.pointerEvents = ''
  document.body.style.overflow = ''
  document.body.removeAttribute('data-scroll-locked')
})
