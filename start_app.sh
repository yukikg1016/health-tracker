#!/bin/bash
# プレビューサーバーから割り当てられた PORT を Streamlit に渡す
PORT=${PORT:-8501}
cd "$(dirname "$0")"

# Mac がスリープしないようにする（アプリ終了時に自動解除）
caffeinate -i &
CAFFEINATE_PID=$!
trap "kill $CAFFEINATE_PID 2>/dev/null" EXIT

exec python3 -m streamlit run health_tracker_app.py \
  --server.port "$PORT" \
  --server.address 0.0.0.0 \
  --server.headless true \
  --server.enableCORS false \
  --server.enableXsrfProtection false
