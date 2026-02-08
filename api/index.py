import sys
import os

# Add project root to sys.path to ensure local modules can be imported
# Current file is api/index.py, so parent of parent is root
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(root_dir)

from main import app
