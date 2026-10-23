#!/bin/zsh
set -euo pipefail

addon_dir="${0:A:h}"
cert_dir="$addon_dir/.certs"
wef_dir="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
launch_agent_dir="$HOME/Library/LaunchAgents"
launch_agent="$launch_agent_dir/com.local.powerpoint-red-marks.plist"
service="gui/$(/usr/bin/id -u)/com.local.powerpoint-red-marks"
for tool in node npm openssl pandoc python3; do
  command -v "$tool" >/dev/null || { echo "缺少 $tool，请先安装。" >&2; exit 1; }
done
node_bin="$(command -v node)"
tool_path="$(dirname "$node_bin"):$(dirname "$(command -v pandoc)"):$(dirname "$(command -v python3)"):/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin"

mkdir -p "$cert_dir" "$wef_dir" "$launch_agent_dir"

if [[ ! -f "$addon_dir/node_modules/mathjax/es5/tex-svg.js" ]]; then
  npm ci --prefix "$addon_dir" --ignore-scripts
fi
"$node_bin" "$addon_dir/build.js"

if [[ ! -f "$cert_dir/localhost.crt" || ! -f "$cert_dir/localhost.key" ]]; then
  openssl req -x509 -newkey rsa:2048 -sha256 -days 825 -nodes \
    -keyout "$cert_dir/localhost.key" \
    -out "$cert_dir/localhost.crt" \
    -subj "/CN=localhost" \
    -addext "subjectAltName=DNS:localhost,IP:127.0.0.1" \
    -addext "extendedKeyUsage=serverAuth"
fi

/usr/bin/security add-trusted-cert -d -r trustRoot \
  -k "$HOME/Library/Keychains/login.keychain-db" \
  "$cert_dir/localhost.crt"

/bin/cp "$addon_dir/manifest.xml" "$wef_dir/red-marks-manifest.xml"
/bin/cp "$addon_dir/com.local.powerpoint-red-marks.plist" "$launch_agent"
/usr/bin/plutil -replace ProgramArguments.0 -string "$node_bin" "$launch_agent"
/usr/bin/plutil -replace ProgramArguments.1 -string "$addon_dir/server.js" "$launch_agent"
/usr/bin/plutil -replace StandardOutPath -string "$addon_dir/server.stdout.log" "$launch_agent"
/usr/bin/plutil -replace StandardErrorPath -string "$addon_dir/server.stderr.log" "$launch_agent"
/usr/bin/plutil -replace EnvironmentVariables.PATH -string "$tool_path" "$launch_agent"
/bin/launchctl bootout "$service" >/dev/null 2>&1 || true
/bin/launchctl bootstrap "gui/$(/usr/bin/id -u)" "$launch_agent"

echo "PowerPoint 快捷工具加载项已安装，本地服务会自动运行。"
echo "确认当前演示文稿已保存后，重新启动 PowerPoint，在“开始”选项卡中启用“PowerPoint 快捷工具”。"
