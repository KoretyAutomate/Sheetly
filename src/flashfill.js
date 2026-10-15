// ---------- Flash Fill: learn a string program from examples.
// A program is a list of parts: literal text, or a piece of a source column (whole value or one
// word token, counted from the start or the end) with a case transform.
const CASES = ['same', 'upper', 'lower', 'proper', 'initial']
const applyCase = (s, k) => (k === 'upper' ? s.toUpperCase() : k === 'lower' ? s.toLowerCase()
  : k === 'proper' ? s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, a, b) => a + b.toUpperCase()) : k === 'initial' ? s.slice(0, 1) : s)
const tokens = (s) => [...s.matchAll(/[\p{L}\p{N}]+/gu)].map((m) => m[0])
function pieces(inputs) {
  const out = []
  inputs.forEach((s, col) => {
    for (const cs of CASES) {
      if (cs !== 'initial') out.push({ col, tok: 'all', cs, text: applyCase(s, cs) })
      const tk = tokens(s)
      tk.forEach((w, k) => {
        out.push({ col, tok: k, cs, text: applyCase(w, cs) })
        out.push({ col, tok: k - tk.length, cs, text: applyCase(w, cs) })
      })
    }
  })
  return out.filter((p) => p.text !== '')
}
function pieceValue(p, inputs) {
  const s = inputs[p.col] ?? ''
  if (p.tok === 'all') return applyCase(s, p.cs)
  const tk = tokens(s), i = p.tok < 0 ? tk.length + p.tok : p.tok
  return i >= 0 && i < tk.length ? applyCase(tk[i], p.cs) : null
}
export function run(program, inputs) {
  let s = ''
  for (const part of program) {
    if (part.lit != null) { s += part.lit; continue }
    const v = pieceValue(part, inputs)
    if (v == null) return null
    s += v
  }
  return s
}
export function learn(examples) {
  const [first, ...rest] = examples
  const ps = pieces(first.inputs)
  const out = first.out
  const candidates = []
  const search = (pos, prog, cost) => {
    if (candidates.length > 4000) return
    if (pos === out.length) { candidates.push({ prog, cost }); return }
    for (const p of ps) if (out.startsWith(p.text, pos)) search(pos + p.text.length, [...prog, { col: p.col, tok: p.tok, cs: p.cs }], cost + 1)
    const last = prog[prog.length - 1]
    const lit = last?.lit != null ? [...prog.slice(0, -1), { lit: last.lit + out[pos] }] : [...prog, { lit: out[pos] }]
    search(pos + 1, lit, cost + 3)
  }
  search(0, [], 0)
  candidates.sort((a, b) => a.cost - b.cost || a.prog.length - b.prog.length)
  // A pure literal is only a pattern if every example agrees on it.
  return candidates.find((c) => rest.every((e) => run(c.prog, e.inputs) === e.out) && (c.prog.some((p) => p.lit == null) || rest.length > 0))?.prog ?? null
}
