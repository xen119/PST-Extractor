module.exports = {
  darkMode: ['class', '[data-theme="dark"]'],
  content: {
    relative: true,
    files: ['./index.html', './src/**/*.{ts,tsx,js,jsx}']
  },
  theme: {
    extend: {
      boxShadow: {
        soft: '0 20px 60px rgba(15, 23, 42, 0.08)'
      },
      borderRadius: {
        xl2: '1.25rem'
      }
    }
  },
  plugins: []
}
