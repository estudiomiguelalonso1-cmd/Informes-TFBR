// Los scripts de la app se cargan como <script> sueltos y comparten UN solo ámbito global.
// Si dos declaran arriba de todo el mismo nombre (`const`, `let`, `function`), el segundo
// tira "Identifier has already been declared" y la página entera queda muerta — sin que
// `node --check` ni `require()` lo noten, porque en Node cada archivo tiene su propio
// ámbito. Ya pasó una vez: motor_tfbr.js declaraba `derivarLayoutSaldos` con un require
// arriba, chocando con la función global de config_tfbr.js.
//
// Correr con: node informe-tfbr/test_scope.js

const fs = require("fs");
const path = require("path");

// El mismo orden en el que index.html los carga.
const ARCHIVOS = [
  "formula_utils.js", "formula_hojas.js", "parser_tfbr.js", "config_tfbr.js",
  "motor_tfbr.js", "periodo_tfbr.js", "validar_tfbr.js", "fixes_tfbr.js",
  "github_tfbr.js", "app.js",
];

// Solo las declaraciones de nivel superior: las que arrancan al principio de la línea.
const RE_DECL = /^(?:const|let|var|function|class)\s+([A-Za-z_$][A-Za-z0-9_$]*)/;

function declaracionesDe(archivo) {
  const texto = fs.readFileSync(path.join(__dirname, archivo), "utf8");
  const nombres = [];
  for (const linea of texto.split("\n")) {
    const m = RE_DECL.exec(linea);
    if (m) nombres.push(m[1]);
  }
  return nombres;
}

const dondeSeDeclara = new Map();
const choques = [];

for (const archivo of ARCHIVOS) {
  for (const nombre of declaracionesDe(archivo)) {
    if (dondeSeDeclara.has(nombre) && dondeSeDeclara.get(nombre) !== archivo) {
      choques.push(`${nombre}: declarado en ${dondeSeDeclara.get(nombre)} y en ${archivo}`);
    } else {
      dondeSeDeclara.set(nombre, archivo);
    }
  }
}

if (choques.length) {
  console.error("✗ Nombres repetidos en el ámbito global (la página no va a cargar):");
  for (const c of choques) console.error("   " + c);
  process.exit(1);
}
console.log(`✓ Sin choques de nombres entre los ${ARCHIVOS.length} scripts (${dondeSeDeclara.size} declaraciones).`);
