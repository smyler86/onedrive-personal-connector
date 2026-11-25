import { createHash } from 'node:crypto';
import { mkdir, readFile, readdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const source = fileURLToPath(new URL('overlay/', import.meta.url));
const target = process.argv[2];
if (!target || process.argv.length !== 3) {
  throw new Error('Usage: node apply-overlay.mjs /path/to/fresh-sites-vinext-checkout');
}
const root = resolve(target);
const starter = new Map(Object.entries({
  'app/layout.tsx': 'fb20da0ae67bb49128afe28c32dba4ff1588e70d6c913a64ef8a0b95f94d63da',
  'app/page.tsx': 'c2293d77b0cff0d149cc153eae96eb6f807f0967b12a177b387f8d9b695f6d7f',
  'db/schema.ts': '13ddacbf0be50e054d450a6c642b7fdc1aa943fd6b99084711e07e58b600390d',
  'drizzle/meta/_journal.json': '90cb374e093894d0efce6f1830d873a42aadd8ba49d2f024495e5cb2e1c8c303',
  'worker/index.ts': 'a0a10229cdda6f5fdb32234e3539007412233bb4938a671866c9295a44a3bdb6',
}));
// The target's deployment identity and installed package metadata belong to its owner.
const readJson = async relative => JSON.parse(await readFile(join(root, relative), 'utf8'));
const [hosting, pkg, lock] = await Promise.all([
  readJson('.openai/hosting.json'), readJson('package.json'), readJson('package-lock.json'),
]);
const miniflare = '4.20260515.0';
if (!pkg.devDependencies?.vinext || lock.packages?.['node_modules/miniflare']?.version !== miniflare ||
    !lock.packages?.['']?.devDependencies) {
  throw new Error('Use the Sites Vinext starter described in README.md. The expected locked test runtime was not found.');
}
if ((hosting.d1 && hosting.d1 !== 'DB') || (hosting.capabilities && !Array.isArray(hosting.capabilities))) {
  throw new Error('This checkout has a different hosting configuration. Use a fresh Sites checkout.');
}
const migrations = await readdir(join(root, 'drizzle'));
if (migrations.some(name => name.endsWith('.sql') && name !== '0000_bumpy_charles_xavier.sql')) {
  throw new Error('The target already has other migrations. Use a fresh Sites checkout.');
}
async function files(folder, prefix = '') {
  const result = [];
  for (const entry of await readdir(folder, { withFileTypes: true })) {
    const relative = prefix + entry.name;
    if (entry.isDirectory()) result.push(...await files(join(folder, entry.name), relative + '/'));
    else if (entry.isFile()) result.push(relative);
    else throw new Error('Unexpected source entry: ' + relative);
  }
  return result;
}
const changes = new Map();
for (const name of await files(source)) {
  const replacement = await readFile(join(source, name));
  let current;
  try {
    current = await readFile(join(root, name));
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  if (current?.equals(replacement)) continue;
  if (current) {
    const normalized = current.toString('utf8').replaceAll('\r\n','\n');
    const hash = createHash('sha256').update(normalized).digest('hex');
    if (starter.get(name) !== hash) throw new Error('Existing source differs at ' + name + '. Use a fresh, compatible Sites checkout.');
  }
  changes.set(name, replacement);
}
hosting.d1 = 'DB';
hosting.capabilities = [...new Set([...(hosting.capabilities ?? []), 'mcp'])];
pkg.devDependencies.miniflare = miniflare;
lock.packages[''].devDependencies.miniflare = miniflare;
for (const [path, document] of [['.openai/hosting.json', hosting], ['package.json', pkg], ['package-lock.json', lock]]) {
  changes.set(path, JSON.stringify(document, null, 2) + '\n');
}
for (const [relative, data] of changes) {
  const path = join(root, relative);
  await mkdir(dirname(path), {recursive: true});
  await writeFile(path, data);
}
console.log('OneDrive source installed. Complete the private deployment steps in README.md.');
