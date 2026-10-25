/** @type {import('tailwindcss').Config} */
// 顏色都接到 src/input.css 的 CSS 變數，深色模式只要換變數值
const v = name => `rgb(var(--c-${name}) / <alpha-value>)`;
const scale = (hue, steps) => Object.fromEntries(steps.map(s => [s, v(`${hue}-${s}`)]));

module.exports = {
  content: [
    './*.html',
    './*.js'
  ],
  theme: {
    extend: {
      borderColor: { DEFAULT: v('slate-200') },
      colors: {
        white: v('white'),
        slate: scale('slate', [50, 100, 200, 300, 400, 500, 600, 700, 800, 900]),
        blue: scale('blue', [50, 500, 600, 700]),
        green: scale('green', [100, 700, 800]),
        red: scale('red', [50, 100, 500, 600, 700]),
        amber: scale('amber', [50, 700]),
      },
    },
  },
  plugins: [
    require('@tailwindcss/forms'),
  ],
}
