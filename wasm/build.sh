#!/bin/bash
set -e

echo "Building WASM module..."

# wasm-pack 설치 확인
if ! command -v wasm-pack &> /dev/null; then
    echo "Installing wasm-pack..."
    curl https://rustwasm.github.io/wasm-pack/installer/init.sh -sSf | sh
fi

# 빌드
wasm-pack build --target web --out-dir ../src/wasm

echo "Build complete! Output in src/wasm/"
