# Deployment Guide - Email Mass Sender

This guide covers deploying the Email Mass Sender application to production environments.

## 🚀 Quick Start

### 1. Production Setup
```bash
# Clone and setup
git clone <repository-url>
cd email-mass-sender
pip install -r requirements.txt

# Copy environment file
cp env.example .env

# Edit configuration
nano .env
```

### 2. Run Application
```bash
# Using the launcher (recommended)
python run.py

# Or directly
python main.py
```

## 🔧 Production Configuration

### Environment Variables
Update your `.env` file with production values:

```env
# Security
SECRET_KEY=your-very-secure-secret-key-32-chars-minimum
DEBUG=False

# Database (PostgreSQL recommended)
DATABASE_URL=postgresql://user:password@localhost:5432/email_sender

# OAuth Redirects (use your domain)
OFFICE365_REDIRECT_URI=https://yourdomain.com/callback/office365
GMAIL_REDIRECT_URI=https://yourdomain.com/callback/gmail
YAHOO_REDIRECT_URI=https://yourdomain.com/callback/yahoo
HOTMAIL_REDIRECT_URI=https://yourdomain.com/callback/hotmail

# Performance
MAX_CONCURRENT_ACCOUNTS=100
MAX_CONCURRENT_PER_PROVIDER=25
RATE_LIMIT_PER_MINUTE=50
```

## 🐳 Docker Deployment

### Dockerfile
```dockerfile
FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

EXPOSE 8080

CMD ["python", "main.py"]
```

### Docker Compose
```yaml
version: '3.8'
services:
  email-sender:
    build: .
    ports:
      - "8080:8080"
    environment:
      - DATABASE_URL=postgresql://postgres:password@db:5432/email_sender
    depends_on:
      - db
    volumes:
      - ./logs:/app/logs
      - ./uploads:/app/uploads

  db:
    image: postgres:13
    environment:
      - POSTGRES_DB=email_sender
      - POSTGRES_USER=postgres
      - POSTGRES_PASSWORD=password
    volumes:
      - postgres_data:/var/lib/postgresql/data

volumes:
  postgres_data:
```

### Deploy with Docker
```bash
# Build and run
docker-compose up -d

# View logs
docker-compose logs -f email-sender
```

## ☁️ Cloud Deployment

### AWS EC2
1. **Launch EC2 Instance**
   - Ubuntu 20.04 LTS
   - t3.medium or larger
   - Security group: HTTP (80), HTTPS (443), SSH (22)

2. **Install Dependencies**
   ```bash
   sudo apt update
   sudo apt install python3 python3-pip nginx postgresql
   ```

3. **Setup Application**
   ```bash
   git clone <repository-url>
   cd email-mass-sender
   pip3 install -r requirements.txt
   ```

4. **Configure Nginx**
   ```nginx
   server {
       listen 80;
       server_name yourdomain.com;
       
       location / {
           proxy_pass http://127.0.0.1:8080;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

5. **Setup SSL with Let's Encrypt**
   ```bash
   sudo apt install certbot python3-certbot-nginx
   sudo certbot --nginx -d yourdomain.com
   ```

### Google Cloud Platform
1. **Create VM Instance**
   - Ubuntu 20.04 LTS
   - e2-medium or larger
   - Allow HTTP/HTTPS traffic

2. **Deploy Application**
   ```bash
   # Install dependencies
   sudo apt update
   sudo apt install python3 python3-pip nginx
   
   # Clone and setup
   git clone <repository-url>
   cd email-mass-sender
   pip3 install -r requirements.txt
   ```

3. **Configure Load Balancer**
   - Create HTTP(S) load balancer
   - Add backend service
   - Configure SSL certificate

### DigitalOcean App Platform
1. **Create App**
   - Source: GitHub repository
   - Build command: `pip install -r requirements.txt`
   - Run command: `python main.py`

2. **Environment Variables**
   - Add all required environment variables
   - Set up database add-on

## 🔒 Security Hardening

### 1. Change Default Credentials
```python
# In main.py, update login credentials
if username == 'your-admin-username' and password == 'your-secure-password':
```

### 2. Enable HTTPS
```env
SECURE_SSL_REDIRECT=True
SESSION_COOKIE_SECURE=True
SESSION_COOKIE_HTTPONLY=True
```

### 3. Database Security
```env
# Use strong database credentials
DATABASE_URL=postgresql://secure_user:strong_password@localhost:5432/email_sender
```

### 4. Firewall Configuration
```bash
# UFW (Ubuntu)
sudo ufw allow 22/tcp
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp
sudo ufw enable
```

### 5. Regular Updates
```bash
# Update system packages
sudo apt update && sudo apt upgrade

# Update Python packages
pip install --upgrade -r requirements.txt
```

## 📊 Monitoring & Logging

### 1. Log Management
```bash
# Rotate logs
sudo logrotate -f /etc/logrotate.d/email-sender

# Monitor logs
tail -f logs/email_sender.log
```

### 2. System Monitoring
```bash
# Install monitoring tools
sudo apt install htop iotop nethogs

# Monitor resources
htop
```

### 3. Application Monitoring
- Set up alerts for high error rates
- Monitor email delivery success rates
- Track account health and token expiration

## 🔄 Backup & Recovery

### 1. Database Backup
```bash
# PostgreSQL backup
pg_dump email_sender > backup_$(date +%Y%m%d).sql

# Restore
psql email_sender < backup_20231201.sql
```

### 2. File Backup
```bash
# Backup uploads and logs
tar -czf backup_$(date +%Y%m%d).tar.gz uploads/ logs/
```

### 3. Automated Backups
```bash
#!/bin/bash
# backup.sh
DATE=$(date +%Y%m%d)
pg_dump email_sender > /backups/db_$DATE.sql
tar -czf /backups/files_$DATE.tar.gz uploads/ logs/
find /backups -name "*.sql" -mtime +30 -delete
find /backups -name "*.tar.gz" -mtime +30 -delete
```

## 🚨 Troubleshooting

### Common Issues

#### Application Won't Start
```bash
# Check Python version
python3 --version

# Check dependencies
pip3 list

# Check logs
tail -f logs/email_sender.log
```

#### OAuth Issues
- Verify redirect URIs match exactly
- Check client ID and secret
- Ensure API permissions are granted

#### Database Connection Issues
```bash
# Test PostgreSQL connection
psql -h localhost -U username -d email_sender

# Check database status
sudo systemctl status postgresql
```

#### High Memory Usage
- Reduce `MAX_CONCURRENT_ACCOUNTS`
- Increase server RAM
- Monitor with `htop`

### Performance Optimization

#### Database Optimization
```sql
-- Add indexes
CREATE INDEX idx_email_accounts_provider ON email_accounts(provider);
CREATE INDEX idx_email_logs_sent_at ON email_logs(sent_at);
```

#### Application Optimization
```env
# Reduce concurrency for stability
MAX_CONCURRENT_ACCOUNTS=50
MAX_CONCURRENT_PER_PROVIDER=10

# Increase timeouts
REQUEST_TIMEOUT=60
RETRY_ATTEMPTS=3
```

## 📈 Scaling

### Horizontal Scaling
1. **Load Balancer**: Distribute traffic across multiple instances
2. **Database**: Use read replicas for reporting
3. **Caching**: Implement Redis for session storage

### Vertical Scaling
1. **CPU**: Increase for more concurrent operations
2. **RAM**: More memory for larger campaigns
3. **Storage**: SSD for better I/O performance

## 🔧 Maintenance

### Daily Tasks
- Monitor error logs
- Check account health
- Review sending statistics

### Weekly Tasks
- Update dependencies
- Clean old logs
- Review security logs

### Monthly Tasks
- Rotate API keys
- Update SSL certificates
- Performance review

---

**🎉 Your Email Mass Sender is now production-ready!**

For support and updates, check the project repository and documentation.
