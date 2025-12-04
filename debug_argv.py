import sys
import os

print(f"sys.argv: {sys.argv}")
print(f"type(sys.argv): {type(sys.argv)}")
print("-" * 20)
print(f"sys.argv: {sys.argv}")
print(f"type(sys.argv): {type(sys.argv)}")
print("-" * 20)

try:
    print("os.path.abspath(sys.argv) を試行します...")
    path = os.path.abspath(sys.argv)
    print(f"成功: {path}")
except TypeError as e:
    print(f"失敗: {e}")

print("-" * 20)

try:
    print("os.path.abspath(sys.argv) を試行します...")
    path = os.path.abspath(sys.argv)
    script_dir = os.path.dirname(path)
    print(f"成功: {script_dir}")
except Exception as e:
    print(f"失敗: {e}")