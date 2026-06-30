import { create } from 'zustand'
import { createJSONStorage, persist } from 'zustand/middleware'

type Theme = 'light' | 'dark'

interface UiState {
  theme: Theme
  sidebarCollapsed: boolean
  previewCollapsed: boolean
  settingsMenuOpen: boolean
  hiddenFiltersOpen: boolean
  setTheme: (theme: Theme) => void
  toggleTheme: () => void
  setSidebarCollapsed: (collapsed: boolean) => void
  toggleSidebarCollapsed: () => void
  setPreviewCollapsed: (collapsed: boolean) => void
  togglePreviewCollapsed: () => void
  setSettingsMenuOpen: (open: boolean) => void
  setHiddenFiltersOpen: (open: boolean) => void
}

const normalizeTheme = (value: unknown): Theme => (value === 'dark' ? 'dark' : 'light')

const memoryStorage = (() => {
  const data = new Map<string, string>()
  return {
    get length() {
      return data.size
    },
    clear() {
      data.clear()
    },
    getItem(key: string) {
      return data.has(key) ? data.get(key)! : null
    },
    key(index: number) {
      return Array.from(data.keys())[index] ?? null
    },
    removeItem(key: string) {
      data.delete(key)
    },
    setItem(key: string, value: string) {
      data.set(key, value)
    }
  } as Storage
})()

function getUiStorage(): Storage {
  try {
    if (typeof window === 'undefined' || typeof window.localStorage === 'undefined') {
      return memoryStorage
    }
    return window.localStorage
  } catch {
    return memoryStorage
  }
}

export const useUiStore = create<UiState>()(
  persist(
    (set, get) => ({
      theme: 'light',
      sidebarCollapsed: false,
      previewCollapsed: false,
      settingsMenuOpen: false,
      hiddenFiltersOpen: false,
      setTheme: (theme) => set({ theme: normalizeTheme(theme) }),
      toggleTheme: () =>
        set({
          theme: get().theme === 'dark' ? 'light' : 'dark'
        }),
      setSidebarCollapsed: (collapsed) => set({ sidebarCollapsed: Boolean(collapsed) }),
      toggleSidebarCollapsed: () =>
        set({
          sidebarCollapsed: !get().sidebarCollapsed
        }),
      setPreviewCollapsed: (collapsed) => set({ previewCollapsed: Boolean(collapsed) }),
      togglePreviewCollapsed: () =>
        set({
          previewCollapsed: !get().previewCollapsed
        }),
      setSettingsMenuOpen: (open) => set({ settingsMenuOpen: Boolean(open) }),
      setHiddenFiltersOpen: (open) => set({ hiddenFiltersOpen: Boolean(open) })
    }),
    {
      name: 'pst-mail-explorer.ui',
      storage: createJSONStorage(() => getUiStorage()),
      partialize: (state) => ({
        theme: state.theme,
        sidebarCollapsed: state.sidebarCollapsed,
        previewCollapsed: state.previewCollapsed
      })
    }
  )
)
