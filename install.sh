#!/bin/sh
# Install an official wolfxl CLI release. Set WOLFXL_VERSION to pin a version.
set -eu

release_base=${WOLFXL_RELEASE_BASE_URL:-https://github.com/SynthGL/wolfxl-oss/releases/download}
release_base=${release_base%/}
context_url=$release_base
fail() {
    printf 'wolfxl installer: %s\nURL: %s\n' "$1" "$context_url" >&2
    exit 1
}
require() {
    command -v "$1" >/dev/null 2>&1 || fail "Required command not found: $1"
}
fetch() {
    context_url=$1
    curl --fail --silent --show-error --location --retry 3 --output "$2" "$1" || fail "Download failed"
}

require curl
require tar
require uname
require mktemp
require install
os=$(uname -s) || fail "Cannot detect operating system"
arch=$(uname -m) || fail "Cannot detect architecture"
case "$arch" in
    x86_64|amd64) arch=x86_64 ;;
    aarch64|arm64) arch=aarch64 ;;
    *) fail "Unsupported architecture: $arch" ;;
esac
case "$os" in
    Linux) target=$arch-unknown-linux-musl ;;
    Darwin) target=$arch-apple-darwin ;;
    *) fail "Unsupported operating system: $os (Windows users: use install.ps1)" ;;
esac
if command -v sha256sum >/dev/null 2>&1; then
    hash_command=sha256sum
elif command -v shasum >/dev/null 2>&1; then
    hash_command=shasum
else
    fail "Required command not found: sha256sum or shasum"
fi

tmp=$(mktemp -d) || fail "Cannot create temporary directory"
trap 'rm -rf "$tmp"' 0
trap 'exit 1' HUP INT TERM

version=${WOLFXL_VERSION:-}
if [ -z "$version" ]; then
    # The repository's /latest endpoint belongs to the Python package, so pick
    # the highest stable wolfxl-cli-v* tag from the release list.
    page=1
    : > "$tmp/candidates"
    while :; do
        api_url="https://api.github.com/repos/SynthGL/wolfxl-oss/releases?per_page=100&page=$page"
        fetch "$api_url" "$tmp/releases.json"
        count=$(grep -o '"tag_name": *"' "$tmp/releases.json" | wc -l | tr -d ' ')
        grep -o '"tag_name": *"wolfxl-cli-v[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*"' "$tmp/releases.json" \
            | sed 's/.*"wolfxl-cli-v\([^"]*\)"$/\1/' >> "$tmp/candidates" || :
        [ "$count" -eq 100 ] || break
        page=$((page + 1))
    done
    version=$(sort -t . -k1,1n -k2,2n -k3,3n "$tmp/candidates" | tail -n 1)
    [ -n "$version" ] || fail "No stable wolfxl-cli release found"
fi
version=${version#wolfxl-cli-v}
version=${version#v}
case "$version" in
    ''|*[!0-9A-Za-z.+-]*) fail "Invalid WOLFXL_VERSION: $version" ;;
esac

asset="wolfxl-cli-$version-$target.tar.gz"
asset_url="$release_base/wolfxl-cli-v$version/$asset"
fetch "$release_base/wolfxl-cli-v$version/SHA256SUMS" "$tmp/SHA256SUMS"
# Match exactly one entry for this asset, never a substring or another target.
expected=
while read -r digest filename extra; do
    if [ "$filename" = "$asset" ] || [ "$filename" = "*$asset" ]; then
        [ -z "$expected" ] && [ -z "$extra" ] || fail "Duplicate or malformed checksum entry for $asset"
        expected=$digest
    fi
done < "$tmp/SHA256SUMS"
case "$expected" in
    ''|*[!0-9a-fA-F]*) fail "Missing or invalid SHA256 checksum for $asset" ;;
esac
[ "${#expected}" -eq 64 ] || fail "Invalid SHA256 checksum length for $asset"
fetch "$asset_url" "$tmp/$asset"
if [ "$hash_command" = sha256sum ]; then
    actual=$(sha256sum "$tmp/$asset") || fail "Cannot calculate SHA256"
else
    actual=$(shasum -a 256 "$tmp/$asset") || fail "Cannot calculate SHA256"
fi
actual=${actual%% *}
expected=$(printf '%s' "$expected" | tr 'A-F' 'a-f') || fail "Cannot normalize SHA256"
[ "$actual" = "$expected" ] || fail "SHA256 mismatch for $asset; nothing was installed"
# Extract only the executable, not arbitrary archive paths.
tar -xzf "$tmp/$asset" -C "$tmp" wolfxl || fail "Cannot extract wolfxl"
[ -f "$tmp/wolfxl" ] && [ ! -L "$tmp/wolfxl" ] || fail "Archive does not contain a regular wolfxl executable"
install_dir=${WOLFXL_INSTALL_DIR:-$HOME/.local/bin}
mkdir -p "$install_dir" || fail "Cannot create installation directory: $install_dir"
install -m 755 "$tmp/wolfxl" "$install_dir/wolfxl" || fail "Cannot install wolfxl to $install_dir"
printf 'Installed wolfxl %s to %s/wolfxl\n' "$version" "$install_dir"
case ":${PATH:-}:" in
    *:"$install_dir":*) ;;
    *) printf 'Add this directory to PATH in your shell profile: %s\n' "$install_dir" ;;
esac
