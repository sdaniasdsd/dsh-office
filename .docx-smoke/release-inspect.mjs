// Read the published release back from the GitHub API, including the asset digest.
import { execFileSync } from 'node:child_process';

const json = execFileSync('gh', ['api', 'repos/sdaniasdsd/dsh-office/releases/tags/v0.7.0'], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
const release = JSON.parse(json);
console.log(`tag        : ${release.tag_name}`);
console.log(`target     : ${release.target_commitish}`);
console.log(`draft      : ${release.draft}   prerelease: ${release.prerelease}`);
console.log(`published  : ${release.published_at}`);
console.log(`html_url   : ${release.html_url}`);
for (const asset of release.assets) {
  console.log(`asset      : ${asset.name}`);
  console.log(`  bytes    : ${asset.size}`);
  console.log(`  state    : ${asset.state}`);
  console.log(`  digest   : ${asset.digest ?? '(not reported by this API version)'}`);
  console.log(`  url      : ${asset.browser_download_url}`);
}
