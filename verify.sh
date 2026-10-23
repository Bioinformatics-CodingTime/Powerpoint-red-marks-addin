#!/bin/zsh
set -euo pipefail

addon_dir="${0:A:h}"

[[ -s "$addon_dir/node_modules/mathjax/es5/tex-svg.js" ]]
node "$addon_dir/build.js" --check

/usr/bin/xmllint --noout "$addon_dir/manifest.xml"
/usr/bin/plutil -lint "$addon_dir/com.local.powerpoint-red-marks.plist"
node --check "$addon_dir/server.js"
node --check "$addon_dir/src/commands.js"
node --check "$addon_dir/src/commands.core.js"
node "$addon_dir/test/commands.test.js"
node "$addon_dir/test/mathjax.test.js"
node "$addon_dir/test/server.test.js"
python3 "$addon_dir/test/formula.test.py"

for icon in arrow-16 arrow-32 arrow-80 box-16 box-32 box-80; do
  [[ -s "$addon_dir/assets/$icon.png" ]]
  magick "$addon_dir/assets/$icon.png" -format "%c" histogram:info:- | /usr/bin/grep -q "#FF0000"
done

echo "本地结构检查通过。"
