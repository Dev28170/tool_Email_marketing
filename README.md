# Email Mass Sender

A professional, enterprise-grade email mass sending application with OAuth2.0 authentication, AI-powered content generation, and multi-provider support.

## 🚀 Features

### Core Functionality
- **Multi-Provider Support**: Office365, Gmail, Yahoo Mail, Hotmail/Outlook.com
- **OAuth2.0 Authentication**: Secure, password-free authentication
- **Mass Email Sending**: Send to 100-200+ recipients simultaneously
- **Multithreaded Processing**: Async/concurrent email sending with rate limiting
- **AI Integration**: OpenAI-powered placeholder replacement for dynamic content
- **Attachment Support**: Small and large file attachments with upload sessions
- **BCC Functionality**: Blind carbon copy support
- **Campaign Management**: Create, manage, and track email campaigns

### Advanced Features
- **Rate Limiting**: Per-provider rate limiting and throttling handling
- **Retry Logic**: Exponential backoff for failed sends
- **Account Management**: Add, remove, and monitor email accounts
- **Real-time Monitoring**: Live dashboard with sending statistics
- **Error Handling**: Comprehensive error logging and recovery
- **Security**: Token refresh, secure storage, and audit trails

## 📋 Requirements

### System Requirements
- Python 3.8 or higher
- 4GB RAM minimum (8GB recommended for large campaigns)
- 1GB free disk space
- Internet connection for OAuth and API calls

### Dependencies
- Flask 2.3.3+
- SQLAlchemy 2.0+
- aiohttp 3.8+
- msal 1.24+ (Microsoft authentication)
- google-auth 2.23+ (Google authentication)
- selenium 4.15+ (Browser automation)
- openai 1.3+ (AI integration)
- loguru 0.7+ (Advanced logging)

## 🛠️ Installation

### 1. Clone the Repository
```bash
git clone <repository-url>
cd email-mass-sender
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Environment Setup
Create a `.env` file in the project root:
```env
# Application Settings
SECRET_KEY=your-secret-key-here
DEBUG=False
HOST=127.0.0.1
PORT=8080

# Database
DATABASE_URL=sqlite:///email_sender.db

# OpenAI Configuration
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

### 4. OAuth2.0 Setup

#### Microsoft Azure (Office365/Hotmail)
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" → "App registrations"
3. Click "New registration"
4. Name: "Email Mass Sender"
5. Redirect URI: `http://localhost:8080/callback/office365`
6. Note down: **Application (client) ID**
7. Go to "Certificates & secrets" → "New client secret"
8. Note down: **Client secret value**
9. Go to "API permissions" → "Add a permission"
10. Select "Microsoft Graph" → "Delegated permissions"
11. Add: `Mail.Send`, `Mail.ReadWrite`, `User.Read`
12. Click "Grant admin consent"

Add to `.env`:
```env
OFFICE365_CLIENT_ID=your-client-id
OFFICE365_CLIENT_SECRET=your-client-secret
OFFICE365_TENANT_ID=your-tenant-id
```

#### Google Cloud Console (Gmail)
1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create new project or select existing
3. Enable "Gmail API"
4. Go to "Credentials" → "Create Credentials" → "OAuth 2.0 Client ID"
5. Application type: "Web application"
6. Authorized redirect URIs: `http://localhost:8080/callback/gmail`
7. Note down: **Client ID** and **Client Secret**

Add to `.env`:
```env
GMAIL_CLIENT_ID=your-client-id
GMAIL_CLIENT_SECRET=your-client-secret
```

#### Yahoo Developer (Yahoo Mail)
1. Go to [Yahoo Developer](https://developer.yahoo.com)
2. Create new application
3. Note down: **Client ID** and **Client Secret**
4. Set redirect URI: `http://localhost:8080/callback/yahoo`

Add to `.env`:
```env
YAHOO_CLIENT_ID=your-client-id
YAHOO_CLIENT_SECRET=your-client-secret
```

### 5. Run the Application
```bash
python main.py
```

The application will be available at `http://localhost:8080`

## 🎯 Usage

### 1. Login
- Default credentials: `admin` / `admin`
- Change these in production!

### 2. Add Email Accounts
1. Go to "Accounts" → "Add Account"
2. Click on your email provider
3. Complete OAuth authentication
4. Account will be added automatically

### 3. Create Campaign
1. Go to "Campaigns" → "New Campaign"
2. Enter campaign name and subject
3. Write email content with placeholders (e.g., `[PROJECT_DETAILS]`)
4. Add recipients
5. Start campaign

### 4. AI Placeholder Replacement
Use placeholders in your email content:
- `[PROJECT_DETAILS]` - AI-generated project updates
- `[CLIENT_NAME]` - Client name
- `[COMPANY_NAME]` - Company name
- `[DEADLINE]` - Project deadline
- `[BUDGET]` - Budget information

The AI will replace these with relevant content based on your prompts.

### 5. Monitor Progress
- View real-time statistics on the dashboard
- Monitor campaign progress
- Check account health and success rates
- Review error logs

## 📊 API Endpoints

### Authentication
- `GET /login` - Login page
- `POST /login` - Authenticate user
- `GET /logout` - Logout user

### Accounts
- `GET /accounts` - List all accounts
- `GET /accounts/add` - Add account page
- `GET /accounts/oauth/<provider>` - OAuth redirect
- `GET /callback/<provider>` - OAuth callback
- `POST /accounts/<id>/test` - Test account
- `DELETE /accounts/<id>` - Delete account

### Campaigns
- `GET /campaigns` - List campaigns
- `GET /campaigns/new` - New campaign page
- `POST /campaigns/create` - Create campaign
- `POST /campaigns/<id>/recipients` - Add recipients
- `POST /campaigns/<id>/send` - Send campaign
- `GET /campaigns/<id>/status` - Campaign status

### AI Integration
- `POST /ai/test` - Test AI connection
- `POST /ai/replace` - Replace placeholders

### Monitoring
- `GET /monitoring` - Monitoring dashboard
- `GET /settings` - Application settings

## 🔧 Configuration

### Environment Variables
All configuration is done through environment variables. See the `.env` example above.

### Database
The application uses SQLite by default. For production, consider PostgreSQL:
```env
DATABASE_URL=postgresql://user:password@localhost/email_sender
```

### Logging
Logs are stored in the `logs/` directory:
- `email_sender.log` - General application logs
- `email_sender_errors.log` - Error logs only
- `email_sender_emails.log` - Email operation logs

## 🚨 Security Considerations

### Production Deployment
1. **Change default credentials** immediately
2. **Use HTTPS** for all OAuth redirects
3. **Set strong SECRET_KEY**
4. **Use environment variables** for sensitive data
5. **Enable database encryption**
6. **Set up proper firewall rules**
7. **Regular security updates**

### OAuth Security
- Tokens are stored encrypted
- Automatic token refresh
- Revocation support
- No password storage

### Rate Limiting
- Per-provider rate limits
- Burst protection
- Automatic retry with backoff
- Throttling detection

## 🐛 Troubleshooting

### Common Issues

#### OAuth Authentication Fails
1. Check redirect URIs match exactly
2. Verify client ID and secret
3. Ensure API permissions are granted
4. Check network connectivity

#### Emails Not Sending
1. Verify account tokens are valid
2. Check rate limits
3. Review error logs
4. Test individual accounts

#### AI Placeholders Not Working
1. Verify OpenAI API key
2. Check API quota
3. Review placeholder syntax
4. Test AI connection

#### Performance Issues
1. Reduce concurrent accounts
2. Increase rate limits
3. Check system resources
4. Review database performance

### Logs
Check the following log files for detailed error information:
- `logs/email_sender.log` - General logs
- `logs/email_sender_errors.log` - Error details
- `logs/email_sender_emails.log` - Email operations

## 📈 Performance Optimization

### Recommended Settings
- **Small campaigns (< 100 emails)**: 10-20 concurrent accounts
- **Medium campaigns (100-500 emails)**: 20-50 concurrent accounts
- **Large campaigns (500+ emails)**: 50-100 concurrent accounts

### System Tuning
- Increase `MAX_CONCURRENT_ACCOUNTS` for more parallelism
- Adjust `RATE_LIMIT_PER_MINUTE` based on provider limits
- Use SSD storage for better I/O performance
- Increase RAM for large attachment handling

## 🤝 Support

### Documentation
- [Setup Guide](docs/setup_guide.md)
- [API Documentation](docs/api_documentation.md)
- [User Manual](docs/user_manual.md)

### Issues
Report issues on the project repository with:
- Error logs
- Configuration details
- Steps to reproduce
- System information

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Microsoft Graph API
- Google Gmail API
- Yahoo Mail API
- OpenAI API
- Flask and SQLAlchemy communities

---

**⚠️ Important**: This application is for legitimate email marketing and communication purposes only. Ensure compliance with CAN-SPAM, GDPR, and other applicable regulations.
