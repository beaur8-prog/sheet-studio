import fs from 'node:fs/promises';
import path from 'node:path';
import {androidPrintFixture} from './fixtures/android-print.mjs';
const target=process.argv[2];if(!target)throw Error('Provide the isolated Android test fixture path');
await fs.mkdir(path.dirname(target),{recursive:true});await fs.writeFile(target,JSON.stringify(await androidPrintFixture()));
