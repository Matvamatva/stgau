import subprocess
import os

sh_filename = "_.sh"

sh_content = """#!/bin/sh

whoami

"""
with open(sh_filename, "w") as f:
    f.write(sh_content)
os.chmod(sh_filename, 0o755)

out2 = subprocess.run(
    ["./" + sh_filename],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    universal_newlines=True
)

out1 = out2.stdout

print("Вывод: " + out1)
print("Ошибка: " + str(out2))
