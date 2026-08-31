#!/usr/bin/env python3
import json,sys
from pathlib import Path
assert Path(sys.argv[1]).is_dir()
assert (Path(sys.argv[1])/'SKILL.md').is_file()
print('pipeline structure ok')
