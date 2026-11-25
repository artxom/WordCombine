#!/bin/zsh
cd "$(dirname "$0")" || exit 1
/Library/Frameworks/Python.framework/Versions/3.14/bin/python3 "$(pwd)/word_combine.py"
