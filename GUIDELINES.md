# Writing guidelines

Every change to this repository follows these rules: app code, tests, scripts,
workflows, comments, documentation and commits. They bring together the
project's own rules and the parts of three outside guides that apply here:

- Vercel's Web Interface Guidelines (accessibility, forms, focus, layout).
- Anthropic's frontend-design guidance (visual direction and interface text).
- Anthropic's webapp-testing guidance (testing in a real browser).

When an existing file already follows a consistent style, match it. These
rules decide only where a file is unclear or inconsistent.

## 1. Project shape

Change these only with Ofek's approval.

- Plain HTML, CSS and JavaScript. No framework, no bundler, no build step and
  no runtime dependency.
- The main page is `index.html`, `styles.css`, `app.js`, `sw.js` and
  `manifest.webmanifest` at the root, with a full copy under `beta/`.
- `app.js` is split into numbered feature sections, listed at the top of the
  file. New code goes in the section it belongs to. Add a new section only for
  a genuinely new feature, and update the list at the top.

## 2. JavaScript

- Modern browser JavaScript: `const` by default, `let` only when the value is
  reassigned, never `var`.
- Compare with `===` and `!==`. The one accepted exception is `x == null`,
  which checks for both `null` and `undefined`.
- Convert typed or imported values with `Number()` and check them with
  `isFinite()` before using them in a calculation.
- Names: `camelCase` for variables and functions, `UPPER_SNAKE_CASE` for
  constants. Names are English and say what the value is, including its unit
  when it has one: `distanceKm`, `energyKwh`, `priceIlsPerKwh`.
- One function does one job. Split a function longer than about 40 lines or
  nested more than three levels deep.
- No magic numbers. Prices, thresholds, sizes and timeouts are named
  constants at the top of their section, with their unit and source.
- Money and energy: keep full precision while calculating and round only for
  display. Numbers are shown through `formatNumber` (Hebrew locale) and dates
  through `formatDate`. Don't build a new format by hand.
- No new globals. Shared state lives in the existing state object.
- Async code uses `async`/`await`. Every network call (OneDrive, Microsoft
  Graph, petrol prices, weather) has `try`/`catch` and shows the user a clear
  Hebrew message when it fails. Never swallow an error silently.
- Put user or file data into the page with `textContent`, or pass it through
  `escapeHtml` first. Use `innerHTML` only for markup written in the code.
- Formatting: 2-space indent, semicolons, double quotes, and lines of at most
  120 characters. Break a long line after a comma or an operator, or split a
  long string into parts joined with `+`.
- No dead code, commented-out code or leftover `console.log` in a commit.

## 3. Documentation in code

Comments are in English. Text the user sees is in Hebrew.

- **File header:** every JavaScript and CSS file starts with a comment saying
  what the file is. `app.js` also lists its numbered sections there.
- **Section header:** each numbered section starts with a banner comment with
  its number and name, then one to three lines on what the section owns.
- **JSDoc:** every function that isn't a trivial one-liner gets a JSDoc
  block, like this:

  ```js
  /**
   * Splits a charging period across calendar months by day count.
   * A public charge is counted in the month of its own date.
   *
   * @param {Period} period - Start and end readings with kWh and cost.
   * @param {number} priceIlsPerKwh - Home electricity price, incl. VAT.
   * @returns {MonthShare[]} One entry per month the period touches.
   */
  ```

- **Data shapes:** the shapes stored in the data file (readings, public
  charges, loan, settings) are described with `@typedef` next to the code that
  loads them.
- **Inline comments explain why, not what.** Use them for a business rule, a
  unit conversion, a workaround and the bug it avoids, or where a number comes
  from (for example `// Regulated tariff from 1.7.2026, incl. VAT`).
- **Keep comments true.** When behaviour changes, update the comments and
  JSDoc that describe it in the same change. A comment that no longer matches
  the code is a bug.
- **TODOs:** a `TODO` comment says why it's there, and Ofek is told about it.

## 4. HTML, CSS and accessibility

These rules come from the Web Interface Guidelines.

- `<html lang="he" dir="rtl">`.
- Use the semantic element for the job: `<button>` for actions, `<a>` for
  navigation, `<label>` for every input, `<table>` for tabular data. Never put
  a click handler on a `<div>` or a `<span>`.
- Icon-only buttons get a Hebrew `aria-label`. Decorative icons and images get
  `aria-hidden="true"` or `alt=""`.
- Inputs get the right `type` and `inputmode` (`decimal` for kWh and prices)
  and a meaningful `name`. Never block paste.
- Messages that change on their own get `role="status"` (or
  `aria-live="polite"`) so screen readers announce them.
- Every interactive element shows a visible focus ring. Never remove
  `outline` without a `:focus-visible` replacement.
- Images get `width` and `height`.
- Colours, spacing and corner radius come only from the custom properties in
  `:root`. Don't add new hard-coded colour values elsewhere.
- Prefer logical properties (`margin-inline-start`, `padding-inline`) so
  right-to-left layout works without overrides.
- Respect the phone's safe areas with `env(safe-area-inset-*)` and never
  disable zoom (`user-scalable=no`, `maximum-scale=1`).
- Destructive actions need a confirmation step, such as the existing
  "click again to delete".
- Never use `transition: all`; list the properties that animate.

## 5. Visual design

This section comes from the frontend-design guidance.

- The visual direction is set: electric blue, and the layout Ofek approved on
  beta. New screens follow it. A change of direction is proposed to Ofek with
  a mock-up first and is never made silently.
- Spend boldness in one place per screen and keep the rest quiet. Remove any
  decoration that doesn't help the user read the numbers.
- Structure carries meaning. Use numbering, dividers and badges only when the
  content really is a sequence or a group.
- Motion is rare and answers the user: opening, expanding, confirming. Don't
  add decorative entrance animations.
- Check every change on a phone-sized screen and on a wide screen.

## 6. Hebrew interface text

- Write short, plain, active Hebrew. A button says exactly what happens: "שמור
  קריאה", not "אישור".
- The same action keeps the same word everywhere: the button, the
  confirmation and the error.
- An error message says what went wrong and what to do next. It doesn't
  apologise.
- An empty screen tells the user what to do to fill it.
- Numbers, units and English names inside Hebrew text must not break the
  right-to-left order. Wrap them in `<bdi>` or a `dir="ltr"` span when needed.

## 7. Security and privacy

- The repository is public. Never commit data files (`zeekr-data.json`,
  `zeekr-test.json`, exports, CSVs), tokens, passwords, real readings, loan
  figures or personal details.
- The Entra client id is public by design. Nothing else secret goes in code.
- Tests use made-up data only.
- Never test writes on the main page. Tests that write use `beta/` and the
  test file.

## 8. Tests

This section comes from the webapp-testing guidance and the project's own
test rules (see `tests/README.md`).

- Every change to a calculation adds or updates a case in
  `tests/unit.test.js`, with a hand-worked expected value and a comment
  showing the arithmetic.
- Every new tab, form or button is covered in `tests/browser.test.js`.
- Browser tests wait for the page to finish loading before they inspect it,
  find elements by role, text or id, and fail on any script error.
- Run the affected tests before pushing. A pull request is merged only after
  the GitHub checks pass.

## 9. Versions, commits and pull requests

- Only the release thread uploads to git. Every other thread sends its change
  to the release thread, which pushes it, opens and merges the pull request
  after Ofek approves, and reports the new version to the version-tracking
  thread.
- Every change goes to `beta/` first. It's copied to the root (the main page)
  only after Ofek approves, together with its tests.
- Versions count in tenths: a major number and one digit after the point. Each
  update adds one tenth (4.0, 4.1 … 4.9), and the update after x.9 moves to the
  next major number with .0 (4.9 is followed by 5.0, never 4.10). The beta
  version is the one that counts; the main page takes the beta number when it
  is promoted.
- Every update raises the version. Change `APP_VERSION` in `app.js` and the
  `?v=` keys in `index.html` together; the header shows the version
  automatically. Raise `CACHE_VERSION` in `sw.js` only to throw away
  everything cached by older versions.
- Record every new version in `CHANGELOG.md`, without personal figures.
- Commit messages are in English and in the imperative. Each has one subject
  line of up to about 72 characters, then a body that says why. Make one
  logical change per commit.
- A pull request says what changed, why, how it was tested, the new version
  number, and whether it's beta or main.

## 10. Before finishing, check

- [ ] The code is in the right section and follows sections 2 to 6.
- [ ] Every new or changed function has up-to-date JSDoc, and no comment is
      out of date.
- [ ] The diff has no data files, secrets, dead code or `console.log`.
- [ ] No line is longer than 120 characters.
- [ ] Tests are added or updated, and they pass.
- [ ] The version is raised, and the change is on beta first.
