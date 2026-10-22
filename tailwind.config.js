/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './index.html',
    './special-fund.html',
    './budget-app.js'
  ],
  theme: {
    extend: {},
  },
  plugins: [
    require('@tailwindcss/forms'),
  ],
}
