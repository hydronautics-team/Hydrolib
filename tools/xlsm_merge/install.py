"""Run once per clone after installing requirements.txt."""
from pathlib import Path
import shlex
import subprocess
import sys

script = Path(__file__).with_name('merge.py').resolve()
driver = ' '.join([shlex.quote(sys.executable), shlex.quote(str(script)), '"%O"', '"%A"', '"%B"', '"%A"'])
subprocess.run(['git', 'config', 'merge.hydrolib-xlsm.name', 'Hydrolib component workbook merge'], check=True)
subprocess.run(['git', 'config', 'merge.hydrolib-xlsm.driver', driver], check=True)
subprocess.run(['git', 'config', 'merge.hydrolib-xlsm.recursive', 'binary'], check=True)
print('Installed for this clone. Re-run if the clone or Python environment moves.')
