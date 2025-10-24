#!/usr/bin/env python3
"""
Email Mass Sender - Application Launcher
Simple launcher script with environment validation and error handling
"""

import os
import sys
import asyncio
from pathlib import Path

def check_python_version():
    """Check if Python version is compatible"""
    if sys.version_info < (3, 8):
        print("❌ Error: Python 3.8 or higher is required")
        print(f"   Current version: {sys.version}")
        sys.exit(1)
    print(f"✅ Python version: {sys.version.split()[0]}")

def check_dependencies():
    """Check if required dependencies are installed"""
    required_packages = [
        ('flask', 'flask'),
        ('sqlalchemy', 'sqlalchemy'),
        ('aiohttp', 'aiohttp'),
        ('msal', 'msal'),
        ('google-auth', 'google.auth'),
        ('openai', 'openai'),
        ('loguru', 'loguru')
    ]
    
    missing_packages = []
    for package_name, import_name in required_packages:
        try:
            __import__(import_name)
        except ImportError:
            missing_packages.append(package_name)
    
    if missing_packages:
        print("❌ Error: Missing required packages:")
        for package in missing_packages:
            print(f"   - {package}")
        print("\n   Install with: pip install -r requirements.txt")
        sys.exit(1)
    print("✅ All required dependencies are installed")

def check_environment():
    """Check environment configuration"""
    env_file = Path('.env')
    if not env_file.exists():
        print("⚠️  Warning: .env file not found")
        print("   Copy env.example to .env and configure your settings")
        print("   cp env.example .env")
        return False
    
    print("✅ Environment file found")
    return True

def create_directories():
    """Create necessary directories"""
    directories = ['logs', 'uploads', 'templates', 'sessions', 'static']
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
    print("✅ Required directories created")

def validate_config():
    """Validate configuration"""
    try:
        from config import config
        validation = config.validate_config()
        
        if validation['warnings']:
            print("⚠️  Configuration warnings:")
            for warning in validation['warnings']:
                print(f"   - {warning}")
        
        if validation['errors']:
            print("❌ Configuration errors:")
            for error in validation['errors']:
                print(f"   - {error}")
            return False
        
        print("✅ Configuration is valid")
        return True
    
    except Exception as e:
        print(f"❌ Configuration error: {e}")
        return False

def main():
    """Main launcher function"""
    print("🚀 Email Mass Sender - Starting Application")
    print("=" * 50)
    
    # Pre-flight checks
    check_python_version()
    check_dependencies()
    check_environment()
    create_directories()
    
    if not validate_config():
        print("\n❌ Configuration validation failed")
        print("   Please fix the configuration issues and try again")
        sys.exit(1)
    
    print("\n✅ All checks passed!")
    print("🌐 Starting web server...")
    print("=" * 50)
    
    # Import and run the main application
    try:
        from main import app, config
        print(f"📱 Application will be available at: http://{config.HOST}:{config.PORT}")
        print("🔐 Default login: admin / admin")
        print("⚠️  Remember to change default credentials in production!")
        print("\n" + "=" * 50)
        
        # Run the application
        app.run(
            host=config.HOST,
            port=config.PORT,
            debug=config.DEBUG
        )
    
    except KeyboardInterrupt:
        print("\n\n👋 Application stopped by user")
    except Exception as e:
        print(f"\n❌ Application error: {e}")
        print("   Check the logs for more details")
        sys.exit(1)

if __name__ == '__main__':
    main()
