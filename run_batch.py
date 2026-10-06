import sys
import os

# Redirect execution to scripts/batch_generate.py
script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts", "batch_generate.py")
if __name__ == "__main__":
    with open(script_path, "r", encoding="utf-8") as f:
        code = f.read()
    exec(compile(code, script_path, "exec"))
