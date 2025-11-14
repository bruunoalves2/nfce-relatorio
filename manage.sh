#!/bin/bash

# Gerenciador da aplicação NFe (start/stop/status/open)

PROJECT_DIR="/Users/brunoalves/Desktop/nfe_relatorio"
VENV_DIR="$PROJECT_DIR/venv"
APP_ENTRY="app/app.py"
PORT=8501
ADDRESS="0.0.0.0"

start() {
  cd "$PROJECT_DIR" || exit 1
  chmod +x run_app.sh 2>/dev/null
  ./run_app.sh
}

start_dev() {
  cd "$PROJECT_DIR" || exit 1
  source "$VENV_DIR/bin/activate"
  streamlit run "$APP_ENTRY" --server.port "$PORT" --server.address "$ADDRESS"
}

stop() {
  pkill -f streamlit 2>/dev/null && echo "Parado." || echo "Nenhum processo Streamlit encontrado."
}

status() {
  if lsof -i :$PORT >/dev/null 2>&1; then
    echo "Rodando na porta $PORT"
    lsof -i :$PORT
  else
    echo "Não está rodando na porta $PORT"
  fi
}

open_app() {
  open "http://localhost:$PORT"
}

case "$1" in
  start)
    start
    ;;
  start-dev|dev)
    start_dev
    ;;
  stop)
    stop
    ;;
  status)
    status
    ;;
  open)
    open_app
    ;;
  *)
    echo "Uso: $0 {start|start-dev|stop|status|open}"
    exit 1
    ;;
esac







