#!/usr/bin/env python3
"""
Cross-platform virtual environment and dependency detection for hwpx-convert.

Finds pypandoc-hwpx and Python in the virtual environment,
regardless of OS (macOS, Linux, Windows) or Python version.
"""

import os
import sys
import glob
import shutil
import subprocess


def _find_venv_root():
    """
    Locate the hwpx_env virtual environment root.
    Search order:
      1. HWPX_VENV_PATH environment variable
      2. hwpx_env/ relative to the skill's scripts/ directory
      3. hwpx_env/ relative to the current working directory
      4. hwpx_env/ in the user's home directory
    """
    # 1. Explicit env var
    env_path = os.environ.get('HWPX_VENV_PATH')
    if env_path and os.path.isdir(env_path):
        return env_path

    # 2. Relative to scripts/ directory (skill bundle)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    skill_root = os.path.dirname(script_dir)
    candidate = os.path.join(skill_root, 'hwpx_env')
    if os.path.isdir(candidate):
        return candidate

    # 3. Relative to CWD
    candidate = os.path.join(os.getcwd(), 'hwpx_env')
    if os.path.isdir(candidate):
        return candidate

    # 4. Home directory
    candidate = os.path.join(os.path.expanduser('~'), 'hwpx_env')
    if os.path.isdir(candidate):
        return candidate

    return None


def _get_bin_dir(venv_root):
    """Return the bin/Scripts directory inside a venv, based on OS."""
    if sys.platform == 'win32':
        return os.path.join(venv_root, 'Scripts')
    return os.path.join(venv_root, 'bin')


def find_venv_python(venv_root=None):
    """
    Find the Python executable inside the virtual environment.
    Auto-detects version (python3.13, python3.12, python3, python).
    """
    if venv_root is None:
        venv_root = _find_venv_root()
    if not venv_root:
        return None

    bin_dir = _get_bin_dir(venv_root)

    # Try versioned Python first, then generic
    if sys.platform == 'win32':
        candidates = ['python.exe', 'python3.exe']
    else:
        # Try specific versions first (descending), then generic
        candidates = sorted(
            glob.glob(os.path.join(bin_dir, 'python3.*')),
            reverse=True
        )
        # Extract just filenames for the glob results
        candidates = [os.path.basename(c) for c in candidates
                      if not c.endswith('.cfg')]
        candidates.extend(['python3', 'python'])

    for name in candidates:
        path = os.path.join(bin_dir, name)
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path

    return None


def find_pypandoc_hwpx(venv_root=None):
    """
    Find the pypandoc-hwpx executable inside the virtual environment.
    """
    if venv_root is None:
        venv_root = _find_venv_root()
    if not venv_root:
        return None

    bin_dir = _get_bin_dir(venv_root)

    if sys.platform == 'win32':
        candidates = ['pypandoc-hwpx.exe', 'pypandoc-hwpx']
    else:
        candidates = ['pypandoc-hwpx']

    for name in candidates:
        path = os.path.join(bin_dir, name)
        if os.path.isfile(path):
            return path

    return None


def check_pandoc():
    """Check if pandoc is installed system-wide."""
    return shutil.which('pandoc') is not None


def check_node():
    """Check if Node.js / npx is available (needed for mermaid-cli)."""
    return shutil.which('npx') is not None


def get_env_info():
    """
    Return a dict with all detected environment paths and statuses.
    Useful for diagnostics.
    """
    venv_root = _find_venv_root()
    return {
        'venv_root': venv_root,
        'venv_python': find_venv_python(venv_root),
        'pypandoc_hwpx': find_pypandoc_hwpx(venv_root),
        'pandoc_installed': check_pandoc(),
        'node_installed': check_node(),
        'platform': sys.platform,
        'python_version': sys.version,
    }


def validate_environment():
    """
    Validate the environment and return (ok, messages).
    ok: True if conversion is possible.
    messages: list of diagnostic strings.
    """
    info = get_env_info()
    messages = []
    ok = True

    if not info['venv_root']:
        ok = False
        messages.append("hwpx_env virtual environment not found.")
        messages.append("Run: python3 -m venv hwpx_env && source hwpx_env/bin/activate && pip install pypandoc-hwpx")
    elif not info['venv_python']:
        ok = False
        messages.append(f"Python not found in {info['venv_root']}")
    elif not info['pypandoc_hwpx']:
        ok = False
        messages.append(f"pypandoc-hwpx not found in {info['venv_root']}")
        messages.append("Run: source hwpx_env/bin/activate && pip install pypandoc-hwpx")

    if not info['pandoc_installed']:
        ok = False
        if sys.platform == 'darwin':
            messages.append("pandoc not installed. Run: brew install pandoc")
        elif sys.platform == 'win32':
            messages.append("pandoc not installed. Run: choco install pandoc  (or download from pandoc.org)")
        else:
            messages.append("pandoc not installed. Run: sudo apt install pandoc  (or your package manager)")

    if not info['node_installed']:
        messages.append("Node.js/npx not found. Mermaid diagrams will use the online API fallback.")

    return ok, messages


if __name__ == '__main__':
    ok, msgs = validate_environment()
    info = get_env_info()

    print("=== hwpx-convert Environment Check ===\n")
    print(f"Platform: {info['platform']}")
    print(f"Python:   {info['python_version'].split()[0]}")
    print(f"Venv:     {info['venv_root'] or 'NOT FOUND'}")
    print(f"Python:   {info['venv_python'] or 'NOT FOUND'}")
    print(f"pypandoc: {info['pypandoc_hwpx'] or 'NOT FOUND'}")
    print(f"pandoc:   {'OK' if info['pandoc_installed'] else 'NOT FOUND'}")
    print(f"Node.js:  {'OK' if info['node_installed'] else 'NOT FOUND (mermaid API fallback)'}")

    print()
    if ok:
        print("All OK! Ready to convert.")
    else:
        print("Issues found:")
        for msg in msgs:
            print(f"  - {msg}")
