import sys
from pathlib import Path
sys.path.insert(0, str(Path('src').resolve()))
import me_engine.curve.generation_agent as g
print(g.GenerationAgent().plan('Global Adhesives & Sealants Market'))
