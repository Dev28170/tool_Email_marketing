# Building Email Mass Sender Executable

This guide will help you create a standalone executable (.exe) file from the Email Mass Sender application.

## Prerequisites

1. **Python 3.8 or higher** installed on your system
2. **All project dependencies** installed (run `pip install -r requirements.txt`)
3. **Windows operating system** (for .exe creation)

## Quick Build (Recommended)

### Option 1: Using the Batch Script
1. Double-click `build.bat`
2. Wait for the build process to complete
3. Your executable will be in `dist/EmailMassSender/`

### Option 2: Using Python Script
1. Open Command Prompt in the project directory
2. Run: `python build_exe.py`
3. Wait for the build process to complete

## Manual Build Process

If you prefer to build manually or need to troubleshoot:

### Step 1: Install PyInstaller
```bash
pip install pyinstaller
```

### Step 2: Install Playwright Browsers
```bash
python -m playwright install chromium
python -m playwright install-deps chromium
```

### Step 3: Create Spec File
The build script automatically creates `EmailMassSender.spec` with the correct configuration.

### Step 4: Build Executable
```bash
python -m PyInstaller --clean EmailMassSender.spec
```

### Step 5: Create Distribution Package
The build script automatically creates a complete distribution package.

## Build Output

After successful build, you'll find:

```
dist/EmailMassSender/
├── EmailMassSender.exe          # Main executable
├── START.bat                    # Easy launcher script
├── README_EXECUTABLE.txt        # Instructions for end users
├── env.example                  # Environment template
├── requirements.txt             # Dependencies list
├── README.md                    # Project documentation
├── DEPLOYMENT.md                # Deployment guide
├── logs/                        # Logs directory
├── uploads/                     # Uploads directory
├── sessions/                    # Sessions directory
└── instance/                    # Instance directory
```

## Distribution

### For End Users
1. Copy the entire `dist/EmailMassSender/` folder to the target machine
2. Run `START.bat` or `EmailMassSender.exe`
3. Configure `.env` file as needed

### File Size
- The executable will be approximately 200-300 MB
- This includes all Python dependencies and Playwright browsers
- No additional Python installation required on target machines

## Troubleshooting

### Common Issues

1. **"Python not found"**
   - Install Python 3.8+ and add it to PATH
   - Restart Command Prompt after installation

2. **"Module not found" errors**
   - Run `pip install -r requirements.txt`
   - Make sure you're in the project directory

3. **Playwright browser issues**
   - Run `python -m playwright install chromium`
   - On Linux: `python -m playwright install-deps chromium`

4. **Build fails with "hiddenimports" errors**
   - The spec file includes most common imports
   - Add missing modules to the `hiddenimports` list in the spec file

5. **Large file size**
   - This is normal for applications with many dependencies
   - Consider using `--onefile` flag for a single file (slower startup)

### Performance Notes

- **Startup time**: 10-30 seconds (normal for PyInstaller executables)
- **Memory usage**: 100-200 MB (includes browser automation)
- **Disk space**: 300-500 MB total

### Security Considerations

- The executable includes all source code
- Consider code obfuscation for production use
- Use proper environment variable management for secrets

## Advanced Configuration

### Custom Icon
Replace `static/favicon.ico` with your custom icon before building.

### One-File Executable
To create a single .exe file (slower startup), modify the spec file:
```python
exe = EXE(
    # ... other parameters ...
    onefile=True,  # Add this line
)
```

### Windowed Mode
To hide the console window, modify the spec file:
```python
exe = EXE(
    # ... other parameters ...
    console=False,  # Change to False
)
```

## Support

If you encounter issues:
1. Check the build logs in the console output
2. Ensure all dependencies are properly installed
3. Try building on a clean Python environment
4. Check PyInstaller documentation for advanced options

## Notes

- The build process may take 5-15 minutes depending on your system
- Antivirus software may flag the executable (false positive)
- The executable is self-contained and doesn't require Python on target machines
- All browser automation features are included in the executable

