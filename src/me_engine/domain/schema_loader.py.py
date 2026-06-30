import json
from pathlib import Path


def load_schema(path: str | Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)