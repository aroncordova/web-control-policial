/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'ui-sans-serif', 'system-ui', 'Segoe UI', 'Arial'],
      },
      colors: {
        ink: '#071016',
        panel: '#0c1a22',
        line: '#12303b',
        emova: '#00a0df',
        subte: '#f6c141',
        alert: '#ff5a5f',
      },
      boxShadow: {
        glow: '0 0 32px rgba(0,160,223,0.18)',
      },
    },
  },
  plugins: [],
};
