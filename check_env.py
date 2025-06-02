import sys
import site
import os

print("=== Python Environment Information ===")
print(f"Python Version: {sys.version}")
print(f"Python Executable: {sys.executable}")
print("\nPython Path:")
for path in sys.path:
    print(f"  {path}")

print("\nSite Packages:")
for path in site.getsitepackages():
    print(f"  {path}")

print("\nPIP Location:")
try:
    import pip
    print(f"  {pip.__file__}")
except ImportError:
    print("  pip not found in current environment")

input("\nPress Enter to exit...") 