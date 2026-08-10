"""ASCII entrypoint for Streamlit launch on Windows batch shells."""

from pathlib import Path
import runpy

if __name__ == "__main__":
	target = Path(__file__).with_name("번역앱.py")
	runpy.run_path(str(target), run_name="__main__")
