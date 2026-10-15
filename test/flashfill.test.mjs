// Unit tests for the Flash Fill learner (src/flashfill.js): node test/flashfill.test.mjs
import { learn, run } from '../src/flashfill.js'
import { writeFileSync } from 'fs'
const cases = [
  ['first name from full name', [[['John Smith'], 'John'], [['Mary Jones'], 'Mary']], ['Ann Lee'], 'Ann'],
  ['last name', [[['John Smith'], 'Smith']], ['Ann Lee'], 'Lee'],
  ['initials', [[['John Smith'], 'JS']], ['Ann Lee'], 'AL'],
  ['last, first', [[['John Smith'], 'Smith, John']], ['Ann Lee'], 'Lee, Ann'],
  ['email user', [[['john.smith@example.com'], 'john']], ['ann.lee@corp.io'], 'ann'],
  ['email domain', [[['john@example.com'], 'example.com'], [['ann@corp.io'], 'corp.io']], ['bob@x.org'], 'x.org'],
  ['upper case', [[['john'], 'JOHN']], ['ann'], 'ANN'],
  ['proper case', [[['jOHN smith'], 'John Smith']], ['ann LEE'], 'Ann Lee'],
  ['combine two columns', [[['John', 'Smith'], 'John Smith']], ['Ann', 'Lee'], 'Ann Lee'],
  ['area code', [[['(425) 555-0100'], '425']], ['(206) 555-0199'], '206'],
  ['date parts', [[['2024-03-15'], '15/03/2024']], ['2025-12-01'], '01/12/2025'],
]
const out = []
for (const [name, examples, input, want] of cases) {
  const prog = learn(examples.map(([inputs, o]) => ({ inputs, out: o })))
  const got = prog ? run(prog, input) : null
  out.push(`${got === want ? 'PASS' : 'FAIL'}  ${name}${got === want ? '' : `  got=${JSON.stringify(got)} want=${JSON.stringify(want)}`}`)
}
const txt = `FLASH FILL ${new Date().toISOString()}\npass ${out.filter((l) => l.startsWith('PASS')).length}/${out.length}\n${out.join('\n')}\n`
console.log(txt); writeFileSync(`test_results/flashfill_${new Date().toISOString().slice(0, 10)}.txt`, txt)
