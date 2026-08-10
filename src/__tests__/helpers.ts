import fs from 'node:fs';
import path from 'node:path';

/** Absolute path of the `fixtures` directory. */
export const fixturesDir = path.join(import.meta.dirname, 'fixtures');

/** Absolute path of a file in the `fixtures` directory. */
export const fixturePath = (name: string) => path.join(fixturesDir, name);

/** Contents of a file in the `fixtures` directory. */
export const readFixture = (name: string) =>
  fs.promises.readFile(fixturePath(name));
