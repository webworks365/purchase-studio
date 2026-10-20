import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import {execFileSync} from 'node:child_process';
const root=new URL('../',import.meta.url),version='aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa';
execFileSync(process.execPath,['scripts/build-release.mjs',version],{cwd:root});
const html=await fs.readFile(new URL('site-release/index.html',root),'utf8');
assert.ok(html.includes(`content="${version}"`));assert.ok(html.includes(`auto-update.js?v=${version}`));
assert.equal(JSON.parse(await fs.readFile(new URL('site-release/version.json',root))).version,version);
for(const name of await fs.readdir(new URL('site-release/',root))){
 if(!/\.m?js$/.test(name))continue;
 const text=await fs.readFile(new URL('site-release/'+name,root),'utf8');
 for(const match of text.matchAll(/(['"])(\.\/[^'"\s]+?\.m?js)(\?v=[a-f0-9]+)?\1/g)){
  assert.equal(match[3],`?v=${version}`,name+' dependency '+match[2]);
  await fs.access(new URL('site-release/'+match[2],root));
 }
}
assert.deepEqual(JSON.parse(await fs.readFile(new URL('site-release/samples.json',root))),{files:[]});
console.log('Release build: versioned HTML/module graph, dependency files and empty public templates passed');
