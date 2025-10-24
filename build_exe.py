#!/usr/bin/env python3
"""
Build script for creating Email Mass Sender executable
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print(f"✅ PyInstaller version: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("❌ PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✅ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("❌ Failed to install PyInstaller")
            return False

def create_spec_file():
    """Create PyInstaller spec file"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('env.example', '.'),
        ('requirements.txt', '.'),
        ('README.md', '.'),
        ('DEPLOYMENT.md', '.'),
    ],
    hiddenimports=[
        'flask',
        'flask_sqlalchemy',
        'flask_login',
        'sqlalchemy',
        'aiohttp',
        'msal',
        'google.auth',
        'google.auth.oauthlib',
        'google.auth.httplib2',
        'selenium',
        'playwright',
        'openai',
        'loguru',
        'email_validator',
        'cryptography',
        'bcrypt',
        'pydantic',
        'python_dotenv',
        'aiofiles',
        'chardet',
        'schedule',
        'pytest',
        'pytest_asyncio',
        'black',
        'flake8',
        'email_service',
        'auth',
        'automation',
        'ai',
        'utils',
        'config',
        'database',
        'main',
        'run',
        'task',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='EmailMassSender',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='static/favicon.ico' if os.path.exists('static/favicon.ico') else None,
)
'''
    
    with open('EmailMassSender.spec', 'w') as f:
        f.write(spec_content)
    print("✅ Created PyInstaller spec file")

def install_playwright_browsers():
    """Install Playwright browsers for the executable"""
    print("🔧 Installing Playwright browsers...")
    try:
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        subprocess.check_call([sys.executable, "-m", "playwright", "install-deps", "chromium"])
        print("✅ Playwright browsers installed")
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Playwright browser installation failed: {e}")
        print("   You may need to install browsers manually after building")

def build_executable():
    """Build the executable using PyInstaller"""
    print("🔨 Building executable...")
    try:
        # Clean previous builds
        if os.path.exists('build'):
            shutil.rmtree('build')
        if os.path.exists('dist'):
            shutil.rmtree('dist')
        
        # Build with PyInstaller
        cmd = [sys.executable, "-m", "PyInstaller", "--clean", "EmailMassSender.spec"]
        subprocess.check_call(cmd)
        
        print("✅ Executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return False

def create_distribution():
    """Create distribution package"""
    print("📦 Creating distribution package...")
    
    dist_dir = Path("dist/EmailMassSender")
    if not dist_dir.exists():
        print("❌ Distribution directory not found")
        return False
    
    # Create necessary directories
    (dist_dir / "logs").mkdir(exist_ok=True)
    (dist_dir / "uploads").mkdir(exist_ok=True)
    (dist_dir / "sessions").mkdir(exist_ok=True)
    (dist_dir / "instance").mkdir(exist_ok=True)
    
    # Copy additional files
    files_to_copy = [
        "env.example",
        "README.md",
        "DEPLOYMENT.md",
        "requirements.txt"
    ]
    
    for file in files_to_copy:
        if os.path.exists(file):
            shutil.copy2(file, dist_dir / file)
    
    # Create startup script
    startup_script = '''@echo off
echo Starting Email Mass Sender...
echo.
echo IMPORTANT: 
echo 1. Copy env.example to .env and configure your settings
echo 2. Make sure you have internet connection
echo 3. Default login: admin / admin
echo.
pause
EmailMassSender.exe
pause
'''
    
    with open(dist_dir / "START.bat", 'w') as f:
        f.write(startup_script)
    
    # Create README for distribution
    dist_readme = '''# Email Mass Sender - Executable Version

## Quick Start

1. **Configure Environment:**
   - Copy `env.example` to `.env`
   - Edit `.env` with your settings

2. **Run the Application:**
   - Double-click `START.bat` or run `EmailMassSender.exe`
   - Open your browser to http://localhost:5000
   - Login with: admin / admin

## Important Notes

- Make sure you have internet connection
- The application will create necessary folders automatically
- Check the logs folder for any issues
- Default port is 5000 (change in .env if needed)

## Troubleshooting

- If the application doesn't start, check Windows Firewall
- Make sure no other application is using port 5000
- Check logs folder for error messages

## Support

For issues and support, check the README.md file.
'''
    
    with open(dist_dir / "README_EXECUTABLE.txt", 'w') as f:
        f.write(dist_readme)
    
    print("✅ Distribution package created")
    print(f"📁 Location: {dist_dir.absolute()}")
    return True

def main():
    """Main build function"""
    print("Email Mass Sender - Executable Builder")
    print("=" * 50)
    
    # Check if we're in the right directory
    if not os.path.exists('main.py'):
        print("❌ Error: main.py not found. Run this script from the project root.")
        sys.exit(1)
    
    # Step 1: Check PyInstaller
    if not check_pyinstaller():
        sys.exit(1)
    
    # Step 2: Create spec file
    create_spec_file()
    
    # Step 3: Install Playwright browsers
    install_playwright_browsers()
    
    # Step 4: Build executable
    if not build_executable():
        sys.exit(1)
    
    # Step 5: Create distribution
    if not create_distribution():
        sys.exit(1)
    
    print("\n" + "=" * 50)
    print("Build completed successfully!")
    print("Your executable is ready in: dist/EmailMassSender/")
    print("Run START.bat to launch the application")
    print("=" * 50)

if __name__ == '__main__':
    main()
