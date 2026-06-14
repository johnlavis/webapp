#!/bin/bash
set -e

REPO="https://github.com/johnlavis/webapp"
DIR="webapp"

if [ -d "$DIR" ]; then
  echo "Updating existing install in ./$DIR ..."
  git -C "$DIR" pull
else
  echo "Cloning $REPO ..."
  git clone "$REPO"
fi

cd "$DIR"

echo "Installing dependencies..."
pip install -r requirements.txt --quiet

echo ""
echo "Done! Start the proxy with:"
echo "  cd $DIR"
echo "  python3 proxy.py start"
