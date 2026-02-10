/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    './index.html',
    './budget-app.js'
  ],
  theme: {
    extend: {},
  },
  plugins: [
    require('@tailwindcss/forms'),
  ],
}
