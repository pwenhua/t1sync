# Project Environment & Python Setup Reference

### 2.1 Python vs Compiler
Python is an interpreted runtime. The VS Code Python extension does not include Python itself; it provides:
- Language features (completion, linting via Pylance/flake8)
- Debugging integration
- Virtual environment detection and interpreter selection
- Test runner integration

You still must install a Python interpreter separately.

### 2.2 Virtual Environment Purpose
A virtual environment isolates dependencies per project. By activating it and pointing VS Code to its interpreter,
you ensure code execution, `pip`, and installed packages do not collide with system-wide or other projects.

## 3. Setup Workflow

### 3.1 Create & Activate Virtual Environment (Windows PowerShell)
```powershell
# create
python -m venv .venv

# activate (first time might require policy change)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\.venv\Scripts\Activate.ps1
```

### 3.2 Upgrade Core Packaging Tools
```powershell
python -m pip install --upgrade pip setuptools wheel
```

### 3.3 Install Dependencies

#### Online
```powershell
python -m pip install -r requirements.txt
```

#### Offline (with pre-downloaded wheels)
Prepare a local directory `wheels/` with all required `.whl`s:
```powershell
python -m pip install --no-index --find-links .\wheels -r requirements.txt
```

### 3.4 Exporting Lock
```powershell
python -m pip freeze > requirements.txt
```

### 3.5 Downloading Wheels on a Connected Machine
```powershell
python -m pip download -r requirements.txt --dest wheels/
```

## 4. VS Code Integration

- Use Command Palette (`Ctrl+Shift+P`) → `Python: Select Interpreter` → select `.venv\Scripts\python.exe`.
- Recommended `.vscode/settings.json` snippet:
```json
{
  "python.analysis.typeCheckingMode": "basic",
  "python.formatting.provider": "black",
  "python.pythonPath": ".venv\\Scripts\\python.exe",
  "python.linting.enabled": true,
  "python.linting.flake8Enabled": true,
  "python.testing.pytestEnabled": true,
  "python.testing.unittestEnabled": false
}
```

## 5. Common Errors & Troubleshooting

### 5.1 `python is not recognized`
- Python executable not on PATH. Either add it or point VS Code directly to the interpreter.

### 5.2 Wrong interpreter used
- Verify with:
  ```powershell
  python -V
  python -m pip list
  ```
- Ensure VS Code status bar shows the expected `.venv` interpreter.
    ```

### 5.5 Activation / Execution Policy issues
- If activation fails in PowerShell:
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

## 6. Rebuild / Migration Recipe

1. Delete previous env:
   ```powershell
| Remove-Item | `Remove-Item -Recurse -Force .\.venv`|
| Create venv | `python -m venv .venv` |
| Activate (PowerShell) | `.\.venv\Scripts\Activate.ps1` |
| Upgrade pip(online) | `python -m pip install --upgrade pip setuptools wheel` |
| Upgrade pip(offline)| `python -m pip install --no-index c:\Temp\wheels\setuptools-80.9.0-py3-none-any.whl`|  
| Upgrade pip(offline)| `python -m pip install --no-index c:\Temp\wheels\wheel-0.45.1-py3-none-any.whl`| 
## download dependencies  
| Export lockfile | `python -m pip freeze > requirements.txt`|
| Download all to wheels | `python -m pip download -r requirements.txt --dest D:\Temp\wheels` |
| Download specific to wheels | `python -m pip download --only-binary=:all:  psycopg2-binary  --dest D:\Temp\wheels   --python-version 313` |

## Install requirements   
| Install requirements (online) | `python -m pip install -r requirements.txt` |
| Install requirements (offline) | `python -m pip install --no-index --find-links C:\temp\wheels -r requirements.txt` |
| Install a specific | `python -m pip install --no-index --find-links="C:\temp\wheels" openpyxl`|
| Install a specific without dependencies  | `python -m pip install C:\Temp\wheels\opencv_python-4.12.0.88-cp37-abi3-win_amd64.whl --no-deps` |
| Install a specific with dependencies|`python -m pip install C:\Temp\wheels\numpy-2.3.3-cp313-cp313-win_amd64.whl --no-index --find-links ` |
| Install a specific | `python -m pip install C:\Temp\wheels\opencv_python-4.12.0.88-cp37-abi3-win_amd64.whl` |


## Uninstall
| normal |  `python -m pip uninstall opencv-python` |
| force | `python -m pip uninstall opencv-python --yes` |
| force-all | `python -m pip uninstall --yes opencv-python` |

| Purge pip cache | `python -m pip cache purge` |

## 8. Git / Project Portability Notes

- Do **not** copy `.venv` between machines. Instead: - Clone repo - install dependencies 

## 9. Suggested File Layout
```
project/
  ├── .venv/                # virtual environment (ignored)
  ├── requirements.txt      # pinned dependencies
  ├── wheels/              # optional offline wheel cache
  ├── .vscode/
  │    └── settings.json    # VS Code recommendations
  └── src/                 # your code
```

## 10. Next Steps / Enhancements
- Automate environment bootstrapping with a `setup.ps1` or Makefile.
- Validate wheel compatibility programmatically (`pip debug --verbose`).
- Add pre-commit hooks to ensure formatting/linting consistency.


## 11. where Python
If you have activated a virtual environment (e.g., with .\.venv\Scripts\Activate.ps1), it uses the Python from .venv\Scripts\python.exe.
If you have not activated a virtual environment, it uses the global Python from your system PATH.

python -c "import sys; print(sys.executable)"
This will show you the full path of the Python interpreter that is currently active.

.venv/Scripts/python.exe is not a full hard copy of your system Python executable. Instead, it is a small launcher (or symlink on Unix) that points to the base Python installation you used to create the virtual environment.

On Windows, .venv/Scripts/python.exe is a copy of the Python launcher, but most of the standard library and binaries are referenced from the base Python installation.
The virtual environment isolates installed packages, but the core Python files are shared unless you use special options (like --copies).
So, .venv/Scripts/python.exe is a lightweight copy/launcher, not a full independent Python installation.