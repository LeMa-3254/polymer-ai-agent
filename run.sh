#!/bin/bash
# Run the Polymer AI Research Agent
# Usage: ./run.sh
# Or with an explicit API key: ANTHROPIC_API_KEY=sk-ant-... ./run.sh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [ -z "$ANTHROPIC_API_KEY" ]; then
  echo "ERROR: ANTHROPIC_API_KEY is not set."
  echo "  Run: export ANTHROPIC_API_KEY=sk-ant-..."
  exit 1
fi

/usr/bin/python3 "$SCRIPT_DIR/agent.py"
