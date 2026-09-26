import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const cliPath = join(__dirname, '..', 'dist', 'cli.js');
const shebang = '#!/usr/bin/env node\n';

try {
  const content = await readFile(cliPath, 'utf8');
  if (!content.startsWith('#!')) {
    await writeFile(cliPath, shebang + content, 'utf8');
    console.log('postbuild: added shebang to dist/cli.js');
  } else {
    console.log('postbuild: shebang already present');
  }
} catch {
  process.exitCode = 0;
  console.log('postbuild: dist/cli.js not found, skipping');
}