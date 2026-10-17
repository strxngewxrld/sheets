import { readFile, writeFile } from 'node:fs/promises';
import { Resvg } from '@resvg/resvg-js';
const svg = await readFile(new URL('../media/icon.svg', import.meta.url));
const png = new Resvg(svg, { fitTo: { mode: 'width', value: 256 } }).render().asPng();
await writeFile(new URL('../media/icon.png', import.meta.url), png);
console.log('media/icon.png generated (256 × 256)');
