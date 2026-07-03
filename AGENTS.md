# Agent Notes

## Project Layout

- Main Python project directory: `SlideMasterOptimizer/`
- Application entry point: `SlideMasterOptimizer/main.py`
- Core OOXML optimizer code: `SlideMasterOptimizer/slidemasteroptimizer/core/`
- PySide6 UI code: `SlideMasterOptimizer/slidemasteroptimizer/ui/`
- Tests: `SlideMasterOptimizer/tests/`

## Python Environment

- This project uses `uv` with a project-local virtual environment at `SlideMasterOptimizer/.venv`.
- Run Python commands from inside `SlideMasterOptimizer/`.
- Prefer activating the Windows batch virtual environment script:

```bat
cd SlideMasterOptimizer
.venv\Scripts\activate.bat
```

- PowerShell activation with `.venv\Scripts\activate.ps1` may fail on this machine because of the local execution policy.

## Dependency Sync

- If `uv sync` fails with the default user cache path, use a project-local cache:

```powershell
cd SlideMasterOptimizer
$env:UV_CACHE_DIR = Join-Path (Get-Location) '.uv-cache'
cmd /c ".venv\Scripts\activate.bat && uv sync"
```

- `.uv-cache/`, `.pytest_cache/`, `.venv/`, and `__pycache__/` are ignored and should not be committed.

## Validation

- Run tests with the project virtual environment active:

```bat
cd SlideMasterOptimizer
.venv\Scripts\activate.bat
python -m pytest
```

- Run a syntax/import smoke check with:

```bat
cd SlideMasterOptimizer
.venv\Scripts\activate.bat
python -m compileall main.py slidemasteroptimizer tests
```

## Application Notes

- The app is a PySide6 client application.
- It accepts a single `.pptx` file by drag and drop or file picker.
- It does not use PowerPoint COM automation; optimization is implemented by direct OOXML package inspection and rewriting.
- The optimizer should not modify the input file in place. Write an optimized copy such as `<name>_optimized.pptx`.
