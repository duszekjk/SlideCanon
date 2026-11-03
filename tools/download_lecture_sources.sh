#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "usage: $0 <lecture_dir>" >&2
  exit 1
fi

lecture_dir="$1"
manifest_file="${2:-$lecture_dir/sources.tsv}"

papers_dir="$lecture_dir/papers"
web_dir="$lecture_dir/web"

if [[ ! -f "$manifest_file" ]]; then
  echo "source manifest not found: $manifest_file" >&2
  exit 1
fi

mkdir -p "$papers_dir" "$web_dir"

download() {
  local url="$1"
  local dest="$2"

  if [[ -f "$dest" ]]; then
    printf 'skip %s\n' "$dest"
    return 0
  fi

  mkdir -p "$(dirname "$dest")"
  printf 'fetch %s\n' "$url"
  curl -fL --retry 3 --retry-delay 2 --connect-timeout 15 \
    -A "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36" \
    -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" \
    -H "Accept-Language: en-US,en;q=0.9" \
    "$url" -o "$dest"
}

while IFS=$'\t' read -r kind url relpath; do
  [[ -z "${kind:-}" ]] && continue
  [[ "${kind:0:1}" == "#" ]] && continue
  if [[ -z "${url:-}" || -z "${relpath:-}" ]]; then
    echo "invalid manifest line: $kind $url $relpath" >&2
    exit 1
  fi
  case "$kind" in
    paper) download "$url" "$papers_dir/$relpath" ;;
    web) download "$url" "$web_dir/$relpath" ;;
    *)
      echo "invalid source kind: $kind" >&2
      exit 1
      ;;
  esac
done < "$manifest_file"

printf 'done %s\n' "$lecture_dir"
