# Setup Guide - Email Mass Sender

This guide will walk you through setting up the Email Mass Sender application from scratch.

## 📋 Prerequisites

### System Requirements
- **Operating System**: Windows 10+, macOS 10.14+, or Linux (Ubuntu 18.04+)
- **Python**: Version 3.8 or higher
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Storage**: 1GB free disk space
- **Network**: Stable internet connection

### Required Accounts
- **Microsoft Azure** account (for Office365/Hotmail)
- **Google Cloud Console** account (for Gmail)
- **Yahoo Developer** account (for Yahoo Mail)
- **OpenAI** account (for AI features)

## 🚀 Step-by-Step Setup

### Step 1: Install Python

#### Windows
1. Download Python from [python.org](https://www.python.org/downloads/)
2. Run the installer
3. **Important**: Check "Add Python to PATH"
4. Verify installation: `python --version`

#### macOS
```bash
# Using Homebrew
brew install python

# Or download from python.org
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt update
sudo apt install python3 python3-pip
```

### Step 2: Clone the Application

```bash
# Clone the repository
git clone <repository-url>
cd email-mass-sender

# Or download and extract ZIP file
```

### Step 3: Install Dependencies

```bash
# Install required packages
pip install -r requirements.txt

# If you encounter permission issues on Linux/macOS:
pip install --user -r requirements.txt
```

### Step 4: Create Environment File

Create a `.env` file in the project root:

```env
# Application Settings
SECRET_KEY=your-very-secure-secret-key-here
DEBUG=False
HOST=127.0.0.1
PORT=8080

# Database
DATABASE_URL=sqlite:///email_sender.db

# OpenAI Configuration (Optional)
OPENAI_API_KEY=sk-your-openai-api-key
OPENAI_MODEL=gpt-4o-mini
OPENAI_MAX_TOKENS=200

# Concurrency Settings
MAX_CONCURRENT_ACCOUNTS=200
MAX_CONCURRENT_PER_PROVIDER=50
REQUEST_TIMEOUT=30
RETRY_ATTEMPTS=5
RETRY_DELAY=1.0

# Rate Limiting
RATE_LIMIT_PER_MINUTE=100
RATE_LIMIT_BURST=10

# File Upload
MAX_ATTACHMENT_SIZE=26214400
UPLOAD_FOLDER=uploads
ALLOWED_EXTENSIONS=txt,pdf,png,jpg,jpeg,gif,doc,docx,xls,xlsx

# Logging
LOG_LEVEL=INFO
LOG_FILE=logs/email_sender.log
```

### Step 5: OAuth2.0 Setup

#### Microsoft Azure (Office365/Hotmail)

1. **Go to Azure Portal**
   - Visit [portal.azure.com](https://portal.azure.com)
   - Sign in with your Microsoft account

2. **Create App Registration**
   - Navigate to "Azure Active Directory"
   - Click "App registrations" → "New registration"
   - Name: "Email Mass Sender"
   - Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: `http://localhost:8080/callback/office365`
   - Click "Register"

3. **Get Credentials**
   - Note down the **Application (client) ID**
   - Go to "Certificates & secrets" → "New client secret"
   - Description: "Email Mass Sender Secret"
   - Expires: "24 months" (recommended)
   - Click "Add"
   - **Important**: Copy the secret value immediately (it won't be shown again)

4. **Configure Permissions**
   - Go to "API permissions" → "Add a permission"
   - Select "Microsoft Graph" → "Delegated permissions"
   - Add the following permissions:
     - `Mail.Send`
     - `Mail.ReadWrite`
     - `User.Read`
   - Click "Grant admin consent" (if you have admin rights)

5. **Add to .env**
   ```env
   OFFICE365_CLIENT_ID=your-application-client-id
   OFFICE365_CLIENT_SECRET=your-client-secret-value
   OFFICE365_TENANT_ID=common
   ```

#### Google Cloud Console (Gmail)

1. **Go to Google Cloud Console**
   - Visit [console.cloud.google.com](https://console.cloud.google.com)
   - Sign in with your Google account

2. **Create Project**
   - Click "Select a project" → "New Project"
   - Name: "Email Mass Sender"
   - Click "Create"

3. **Enable Gmail API**
   - Go to "APIs & Services" → "Library"
   - Search for "Gmail API"
   - Click on it and press "Enable"

4. **Create OAuth Credentials**
   - Go to "APIs & Services" → "Credentials"
   - Click "Create Credentials" → "OAuth 2.0 Client ID"
   - Application type: "Web application"
   - Name: "Email Mass Sender"
   - Authorized redirect URIs: `http://localhost:8080/callback/gmail`
   - Click "Create"

5. **Configure OAuth Consent Screen**
   - Go to "APIs & Services" → "OAuth consent screen"
   - User Type: "External" (unless you have Google Workspace)
   - Fill in required fields:
     - App name: "Email Mass Sender"
     - User support email: Your email
     - Developer contact: Your email
   - Add scopes:
     - `https://www.googleapis.com/auth/gmail.send`
     - `https://www.googleapis.com/auth/gmail.compose`
     - `https://www.googleapis.com/auth/userinfo.email`

6. **Add to .env**
   ```env
   GMAIL_CLIENT_ID=your-client-id
   GMAIL_CLIENT_SECRET=your-client-secret
   ```

#### Yahoo Developer (Yahoo Mail)

1. **Go to Yahoo Developer**
   - Visit [developer.yahoo.com](https://developer.yahoo.com)
   - Sign in with your Yahoo account

2. **Create Application**
   - Go to "My Apps" → "Create an App"
   - App name: "Email Mass Sender"
   - App type: "Web Application"
   - Redirect URI: `http://localhost:8080/callback/yahoo`
   - Permissions: `mail-w`, `mail-r`

3. **Add to .env**
   ```env
   YAHOO_CLIENT_ID=your-client-id
   YAHOO_CLIENT_SECRET=your-client-secret
   ```

### Step 6: OpenAI Setup (Optional)

1. **Create OpenAI Account**
   - Visit [platform.openai.com](https://platform.openai.com)
   - Sign up for an account

2. **Get API Key**
   - Go to "API Keys" → "Create new secret key"
   - Name: "Email Mass Sender"
   - Copy the key (starts with `sk-`)

3. **Add to .env**
   ```env
   OPENAI_API_KEY=sk-your-openai-api-key
   ```

### Step 7: Run the Application

```bash
# Start the application
python main.py
```

You should see output like:
```
 * Running on http://127.0.0.1:8080
 * Debug mode: off
```

### Step 8: Access the Application

1. Open your web browser
2. Go to `http://localhost:8080`
3. Login with default credentials:
   - Username: `admin`
   - Password: `admin`

**⚠️ Important**: Change these credentials immediately in production!

## 🔧 Initial Configuration

### 1. Change Default Password
1. Go to Settings
2. Update admin credentials
3. Save changes

### 2. Test OAuth Connections
1. Go to Accounts → Add Account
2. Try connecting each provider
3. Verify authentication works

### 3. Test AI Integration
1. Go to Settings
2. Click "Test AI Connection"
3. Verify OpenAI integration works

### 4. Create Test Campaign
1. Go to Campaigns → New Campaign
2. Create a simple test campaign
3. Send to a test email address
4. Verify email delivery

## 🚨 Troubleshooting

### Common Issues

#### "Module not found" errors
```bash
# Reinstall dependencies
pip install -r requirements.txt --force-reinstall
```

#### OAuth redirect errors
- Check redirect URIs match exactly
- Verify client ID and secret
- Ensure API permissions are granted

#### Database errors
```bash
# Delete database file and restart
rm email_sender.db
python main.py
```

#### Port already in use
```bash
# Change port in .env file
PORT=8081
```

### Getting Help

1. **Check logs**: Look in `logs/email_sender.log`
2. **Verify configuration**: Ensure all .env variables are set
3. **Test connectivity**: Verify internet connection
4. **Check permissions**: Ensure file system permissions are correct

## 🎯 Next Steps

After successful setup:

1. **Add your email accounts** through the web interface
2. **Create your first campaign** with test recipients
3. **Configure AI placeholders** for dynamic content
4. **Set up monitoring** to track performance
5. **Review security settings** for production use

## 📚 Additional Resources

- [API Documentation](api_documentation.md)
- [User Manual](user_manual.md)
- [Security Guide](security_guide.md)
- [Performance Tuning](performance_guide.md)

---

**🎉 Congratulations!** Your Email Mass Sender is now ready to use. Start with small test campaigns and gradually scale up as you become familiar with the system.
