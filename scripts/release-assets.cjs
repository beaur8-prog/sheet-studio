'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),{createHash}=require('node:crypto');
const {version}=require('../package.json');
async function main(){
 const platform=process.env.TARGET_PLATFORM,arch=process.env.TARGET_ARCH;
 if(!['mac','win','linux'].includes(platform)||!['arm64','x64'].includes(arch))throw Error('Invalid release target');
 const source=path.resolve('release',version,platform+'-'+arch),destination=path.resolve('release-assets');
 await fs.mkdir(destination,{recursive:true});
 const names=(await fs.readdir(source)).filter(name=>/\.(dmg|zip|exe|deb)$/.test(name));
 if(names.length!==(platform==='mac'?2:1))throw Error('Missing release packages');
 const sums=[];
 for(const name of names){const bytes=await fs.readFile(path.join(source,name));await fs.writeFile(path.join(destination,name),bytes);sums.push(createHash('sha256').update(bytes).digest('hex')+'  '+name);}
 for(const name of ['artifact-manifest','runtime-provenance'])await fs.copyFile(path.join(source,name+'.json'),path.join(destination,`${name}-${platform}-${arch}.json`));
 await fs.writeFile(path.join(destination,`SHA256SUMS-${platform}-${arch}.txt`),sums.join('\n')+'\n');
}
main().catch(error=>{console.error(error);process.exitCode=1;});
