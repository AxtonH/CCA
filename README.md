# 📧 Odoo Invoice Follow-Up Manager

A Streamlit application for managing overdue invoice follow-ups with automated email sending capabilities.

## 🚀 Features

- **Odoo Integration**: Connect to Odoo ERP system to fetch overdue invoices
- **Dashboard**: Visual overview of overdue invoices categorized by severity
- **Bulk Email Sender**: Send personalized follow-up emails to multiple clients
- **PDF Generation**: Automatic generation of invoice PDFs for each client
- **Template System**: Configurable email templates based on overdue periods
- **IBAN Attachments**: Automatic attachment of company-specific IBAN letters
- **Demo Mode**: Test the application with sample data

## 📋 Prerequisites

- Python 3.8 or higher
- Odoo instance with API access
- Gmail account for sending emails (or other SMTP provider)
- Google OAuth credentials (for Gmail API)

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/odoo-invoice-followup.git
   cd odoo-invoice-followup
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up secrets (recommended for Streamlit Cloud)**
   ```bash
   cp .streamlit/secrets.toml.example .streamlit/secrets.toml
   ```
   
   Edit `.streamlit/secrets.toml` with your actual credentials:
   ```toml
   # Odoo Configuration
   ODOO_URL = "https://your-odoo-instance.com"
   ODOO_DB = "your-database-name"
   ODOO_USERNAME = "your-username"
   ODOO_PASSWORD = "your-password"

   # Email Configuration
   EMAIL = "your-email@example.com"
   EMAIL_PASSWORD = "your-app-password"

   # Google OAuth Configuration (if needed)
   GOOGLE_CLIENT_ID = "your-google-client-id"
   GOOGLE_CLIENT_SECRET = "your-google-client-secret"
   GOOGLE_PROJECT_ID = "your-google-project-id"
   ```

   **Alternative: Environment Variables**
   ```bash
   cp config.env.example .env
   ```
   
   Edit `.env` with your actual credentials:
   ```env
   # Google OAuth Configuration
   GOOGLE_CLIENT_ID=your-google-client-id
   GOOGLE_CLIENT_SECRET=your-google-client-secret
   GOOGLE_PROJECT_ID=your-google-project-id

   # Odoo Configuration
   ODOO_URL=https://your-odoo-instance.com
   ODOO_DATABASE=your-database-name
   ODOO_USERNAME=your-username
   ODOO_PASSWORD=your-password

   # Email Configuration
   SENDER_PASSWORD=your-email-password
   ```

## 🚀 Running the Application

### Local Development
```bash
streamlit run app.py
```

### Using Scripts
- **Windows**: `run_app.bat`
- **Linux/Mac**: `./run_app.sh`

## 🌐 Deployment

### Streamlit Cloud
1. Push your code to GitHub
2. Connect your repository to [Streamlit Cloud](https://streamlit.io/cloud)
3. Set secrets in Streamlit Cloud dashboard:
   - Go to your app settings
   - Add the following secrets:
     - `ODOO_URL`
     - `ODOO_DB`
     - `ODOO_USERNAME`
     - `ODOO_PASSWORD`
     - `EMAIL`
     - `EMAIL_PASSWORD`
     - `GOOGLE_CLIENT_ID` (if using Gmail API)
     - `GOOGLE_CLIENT_SECRET` (if using Gmail API)
     - `GOOGLE_PROJECT_ID` (if using Gmail API)
4. Deploy!

### Other Platforms
The application can be deployed on any platform that supports Streamlit:
- Heroku
- Railway
- DigitalOcean App Platform
- AWS/GCP/Azure

## 📁 Project Structure

```
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── demo_data.py          # Demo data generator
├── email_templates.py    # Email template system
├── google_oauth_config.py # Google OAuth configuration
├── config.env.example    # Environment variables template
├── run_app.sh           # Shell script for running
├── run_app.bat          # Windows batch script
├── IBAN Letter _ Prezlab FZ LLC .pdf
├── IBAN Letter _ Prezlab Advanced Design Company .pdf
├── .streamlit/
│   └── config.toml      # Streamlit configuration
└── README.md            # This file
```

## 🔧 Configuration

### Email Templates
Templates are defined in `email_templates.py` and support:
- **Initial**: First reminder (30 days payment term)
- **Second**: Follow-up reminder (15 days payment term)
- **Final**: Final notice (7 days payment term)

### IBAN Letters
Automatic attachment of company-specific IBAN letters:
- Prezlab FZ LLC
- Prezlab Advanced Design Company

### Client Classification
Clients are automatically classified based on overdue days:
- **Recent**: ≤ 15 days
- **Moderate**: 16-30 days
- **Severe**: 31+ days

## 🔒 Security Notes

- Never commit `.env` files or `token.pickle` to version control
- Use environment variables for sensitive credentials
- The `.gitignore` file is configured to exclude sensitive files

## 🐛 Troubleshooting

### Common Issues

1. **Odoo Connection Failed**
   - Verify Odoo URL, database, username, and password
   - Check if Odoo instance allows external API access

2. **Email Sending Failed**
   - Verify Gmail credentials and 2FA settings
   - Check if "Less secure app access" is enabled (if using password)
   - Ensure Google OAuth credentials are properly configured

3. **PDF Generation Failed**
   - Check Odoo user permissions for report generation
   - Verify invoice data exists for the client

### Demo Mode
If you can't connect to Odoo, enable Demo Mode in the sidebar to test the application with sample data.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📞 Support

For support and questions, please open an issue on GitHub or contact the development team. 