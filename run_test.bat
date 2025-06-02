@echo off
echo Running Python test...
python test_arch.py > output.txt 2>&1
type output.txt
pause 