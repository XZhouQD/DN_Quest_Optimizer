"""Convenience: python generate_templates.py [out_dir]"""
import sys
from src.templates import main

if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "templates"
    main(out)
