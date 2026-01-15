import subprocess
import sys
import os

req_file = os.path.join(os.path.dirname(__file__), 'requirements.txt')

if not os.path.exists(req_file):
    print('requirements.txt not found!')
    sys.exit(1)

try:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', req_file])
    print('All dependencies installed successfully.')
except subprocess.CalledProcessError as e:
    print('Failed to install dependencies:', e)
    sys.exit(1)
