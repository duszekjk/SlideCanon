#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "usage: $0 <lecture_dir>" >&2
  exit 1
fi

lecture_dir="$1"
batch_size="${BATCH_SIZE:-2}"
repo_root="$(cd "$(dirname "$0")/.." && pwd)"
download_script="${DOWNLOAD_SCRIPT:-$repo_root/tools/download_lecture_sources.sh}"
picaxe_home="${PICAXE_HOME:-/private/tmp/PicAxe/PicAxe-OCR}"
picaxe_python="${PICAXE_PYTHON:-/private/tmp/picaxe-venv/bin/python}"

papers_dir="$lecture_dir/papers"
literature_dir="$lecture_dir/literature"
bundle_dir="$literature_dir/picaxe-out"
batch_root="$bundle_dir/batches"
merged_images="$bundle_dir/extracted_images"
merged_tables="$bundle_dir/tables"
merged_markdown="$bundle_dir/markdown"
flat_images="$literature_dir/extracted_images"
flat_tables="$literature_dir/tables"
flat_markdown="$literature_dir/markdown"
flat_papers="$literature_dir/papers"
flat_web="$literature_dir/web"

if [[ ! -x "$download_script" ]]; then
  echo "download helper not found: $download_script" >&2
  exit 1
fi

if [[ ! -x "$picaxe_python" ]]; then
  echo "PicAxe python not found: $picaxe_python" >&2
  exit 1
fi

if [[ ! -d "$picaxe_home" ]]; then
  echo "PicAxe repo not found: $picaxe_home" >&2
  exit 1
fi

mkdir -p "$bundle_dir" "$batch_root" "$merged_images" "$merged_tables" "$merged_markdown" "$flat_images" "$flat_tables" "$flat_markdown" "$flat_papers" "$flat_web"

"$download_script" "$lecture_dir"

if [[ -d "$papers_dir" ]]; then
  find "$papers_dir" -maxdepth 1 -type f -name '*.pdf' -exec cp -f {} "$flat_papers/" \;
fi

if [[ -d "$lecture_dir/web" ]]; then
  find "$lecture_dir/web" -maxdepth 1 -type f \( -name '*.html' -o -name '*.htm' \) -exec cp -f {} "$flat_web/" \;
fi

pdfs=()
while IFS= read -r pdf; do
  pdfs+=("$pdf")
done < <(find "$papers_dir" -maxdepth 1 -type f -name '*.pdf' | LC_ALL=C sort)

if [[ ${#pdfs[@]} -eq 0 ]]; then
  echo "no PDFs found in $papers_dir" >&2
  exit 1
fi

batch_index=0
total_pdfs="${#pdfs[@]}"
total_batches=$(((total_pdfs + batch_size - 1) / batch_size))
start_ts=$(date +%s)
echo "source prep: $total_pdfs pdfs, batch size $batch_size, $total_batches batches"
for ((i = 0; i < ${#pdfs[@]}; i += batch_size)); do
  batch_index=$((batch_index + 1))
  batch_start=$((i + 1))
  batch_end=$((i + batch_size))
  if [[ $batch_end -gt $total_pdfs ]]; then
    batch_end=$total_pdfs
  fi
  batch_pct=$((100 * batch_end / total_pdfs))
  batch_dir="$(mktemp -d "${TMPDIR:-/private/tmp}/lecture_batch_${batch_index}.XXXXXX")"
  batch_output="$batch_root/batch_$(printf '%03d' "$batch_index")"
  mkdir -p "$batch_output"
  batch_start_ts=$(date +%s)

  echo "source prep: batch $batch_index/$total_batches preparing PDFs $batch_start-$batch_end ($batch_pct% of PDFs)"
  for pdf in "${pdfs[@]:i:batch_size}"; do
    cp "$pdf" "$batch_dir/"
  done

  echo "source prep: batch $batch_index/$total_batches running PicAxe -> $(basename "$batch_output")"
  (
    cd "$picaxe_home"
    "$picaxe_python" main.py --input-dir "$batch_dir" --output-dir "$batch_output"
  )
  echo "source prep: batch $batch_index/$total_batches PicAxe finished ($batch_pct% of PDFs)"

  if [[ -d "$batch_output/extracted_images" ]]; then
    echo "source prep: batch $batch_index/$total_batches merging extracted images"
    find "$batch_output/extracted_images" -type f -name '*.png' -exec cp -f {} "$merged_images/" \; -exec cp -f {} "$flat_images/" \;
  fi

  if [[ -d "$batch_output/tables" ]]; then
    echo "source prep: batch $batch_index/$total_batches merging tables"
    find "$batch_output/tables" -type f -name '*.png' -exec cp -f {} "$merged_tables/" \; -exec cp -f {} "$flat_tables/" \;
  fi

  echo "source prep: batch $batch_index/$total_batches merging markdown"
  find "$batch_output" -type f -name '*.md' -exec cp -f {} "$merged_markdown/" \; -exec cp -f {} "$flat_markdown/" \;

  rm -rf "$batch_dir"
  batch_end_ts=$(date +%s)
  batch_elapsed=$((batch_end_ts - batch_start_ts))
  batches_done=$batch_index
  batches_left=$((total_batches - batches_done))
  if [[ $batches_done -gt 0 ]]; then
    avg_batch=$(((batch_end_ts - start_ts) / batches_done))
  else
    avg_batch=0
  fi
  eta_seconds=$((avg_batch * batches_left))
  eta_min=$((eta_seconds / 60))
  eta_sec=$((eta_seconds % 60))
  echo "source prep: batch $batch_index/$total_batches done in ${batch_elapsed}s ($batch_pct% of PDFs, ETA ${eta_min}m ${eta_sec}s)"
done

echo "done $bundle_dir"
