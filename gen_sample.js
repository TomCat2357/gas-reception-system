const fs = require('fs');
const vm = require('vm');

function loadScriptToContext(file) {
  const code = fs.readFileSync(file, 'utf8');
  const ctx = {
    console,
    setTimeout,
    clearTimeout,
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx, { filename: file });
  return ctx;
}

function main() {
  const csvPath = 'template_example.csv';
  const outPath = 'template_example.generated.html';
  const csv = fs.readFileSync(csvPath, 'utf8');

  const ctx = loadScriptToContext('csv_form_builder.js');
  if (typeof ctx.generateFormFromCsv !== 'function') {
    throw new Error('generateFormFromCsv not found after loading csv_form_builder.js');
  }
  const res = ctx.generateFormFromCsv(csv);
  if (!res || !res.success) {
    throw new Error('Failed to generate HTML: ' + (res && res.error ? res.error : 'unknown error'));
  }
  fs.writeFileSync(outPath, res.html, 'utf8');
  console.log('Wrote', outPath, 'size=', Buffer.byteLength(res.html, 'utf8'));
}

main();

