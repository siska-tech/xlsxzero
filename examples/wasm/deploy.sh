#!/bin/bash
# WASM サンプルのデプロイ用スクリプト

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

echo "Building WASM module..."
cd "$SCRIPT_DIR"
wasm-pack build --target web --out-dir pkg --release

echo "Build completed successfully!"
echo ""
echo "To deploy:"
echo "1. Copy the following files to your web server:"
echo "   - index.html"
echo "   - index.js"
echo "   - pkg/ (entire directory)"
echo ""
echo "2. Or use one of the following services:"
echo "   - GitHub Pages: See README.md for instructions"
echo "   - Netlify: Drag and drop the examples/wasm directory"
echo "   - Vercel: Import the repository and set build settings"
echo ""
echo "3. For GitHub Pages, you can also use:"
echo "   mkdir -p $PROJECT_ROOT/docs/wasm"
echo "   cp index.html index.js $PROJECT_ROOT/docs/wasm/"
echo "   cp -r pkg $PROJECT_ROOT/docs/wasm/"

