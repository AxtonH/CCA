import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import json
import os
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import tempfile
import base64
import time
from streamlit_tags import st_tags
import re
import io
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import tempfile
import os
import time

# Load environment variables
load_dotenv()

# Load secrets from Streamlit secrets.toml
def get_secret(key, default=None):
    """Get secret from Streamlit secrets or environment variables"""
    try:
        return st.secrets[key]
    except:
        return os.getenv(key, default)

# Page configuration
st.set_page_config(
    page_title="Odoo Invoice Follow-Up Manager",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sidebar-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1f77b4;
        margin-bottom: 1rem;
    }
    .warning-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .invoice-table {
        background-color: #f8f9fa;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .email-preview {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .dataframe {
        font-size: 14px;
        line-height: 1.4;
    }
    .dataframe th {
        background-color: #f0f2f6;
        font-weight: bold;
        padding: 12px 8px;
        text-align: left;
        border-bottom: 2px solid #dee2e6;
    }
    .dataframe td {
        padding: 10px 8px;
        border-bottom: 1px solid #e9ecef;
    }
    .dataframe tr:hover {
        background-color: #f8f9fa;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'odoo_connected' not in st.session_state:
    st.session_state.odoo_connected = False
if 'overdue_invoices' not in st.session_state:
    st.session_state.overdue_invoices = []
if 'clients_missing_email' not in st.session_state:
    st.session_state.clients_missing_email = []

class OdooConnector:
    def __init__(self, url, database, username, password):
        self.url = url
        self.database = database
        self.username = username
        self.password = password
        self.uid = None
        self.models = None
    
    def connect(self):
        """Connect to Odoo and authenticate with timeout"""
        try:
            import xmlrpc.client
            import socket
            
            # Set timeout for connections
            socket.setdefaulttimeout(30)  # 30 seconds timeout
            
            # Connect to Odoo
            common = xmlrpc.client.ServerProxy(f'{self.url}/xmlrpc/2/common')
            self.uid = common.authenticate(self.database, self.username, self.password, {})
            
            if self.uid:
                self.models = xmlrpc.client.ServerProxy(f'{self.url}/xmlrpc/2/object')
                return True
            return False
        except Exception as e:
            st.error(f"Connection failed: {str(e)}")
            return False
    
    def get_overdue_invoices(self, progress_callback=None):
        """Fetch overdue invoices from Odoo with batch loading for better performance"""
        try:
            if not self.models:
                return []
            
            # Get current date
            today = datetime.now().date()
            
            if progress_callback:
                progress_callback("Searching for overdue invoices...", 0.1)
            
            # Search for overdue invoices
            invoice_ids = self.models.execute_kw(
                self.database, self.uid, self.password,
                'account.move', 'search',
                [[('state', '=', 'posted'), 
                  ('move_type', '=', 'out_invoice'),
                  ('payment_state', '!=', 'paid'),
                  ('invoice_date_due', '<', today.isoformat())]]
            )
            
            if not invoice_ids:
                return []
            
            if progress_callback:
                progress_callback(f"Found {len(invoice_ids)} overdue invoice IDs.", 0.2)
            
            # Batch fetch all invoices at once (with limit for performance)
            batch_size = 1000  # Limit batch size to prevent server overload
            invoices_data = []
            
            if progress_callback:
                progress_callback(f"Fetching {len(invoice_ids)} invoices...", 0.3)
            
            for i in range(0, len(invoice_ids), batch_size):
                batch_ids = invoice_ids[i:i + batch_size]
                batch_data = self.models.execute_kw(
                    self.database, self.uid, self.password,
                    'account.move', 'read',
                    [batch_ids],  # wrap in a list
                    {'fields': ['name', 'invoice_date_due', 'amount_residual', 'partner_id', 'invoice_origin', 'currency_id', 'company_id']}
                )
                invoices_data.extend(batch_data)
                
                # Update progress
                if progress_callback:
                    progress = 0.3 + (i / len(invoice_ids)) * 0.3
                    progress_callback(f"Fetched {min(i + batch_size, len(invoice_ids))}/{len(invoice_ids)} invoices...", progress)
            
            # Extract unique partner IDs, currency IDs, and company IDs
            partner_ids = list(set([inv['partner_id'][0] for inv in invoices_data if inv.get('partner_id')]))
            currency_ids = list(set([inv['currency_id'][0] for inv in invoices_data if inv.get('currency_id')]))
            company_ids = list(set([inv['company_id'][0] for inv in invoices_data if inv.get('company_id')]))
            
            # Batch fetch all partners at once (with limit for performance)
            partners_data = {}
            if partner_ids:
                if progress_callback:
                    progress_callback(f"Fetching {len(partner_ids)} client records...", 0.7)
                
                partners_list = []
                for i in range(0, len(partner_ids), batch_size):
                    batch_partner_ids = partner_ids[i:i + batch_size]
                    batch_partners = self.models.execute_kw(
                        self.database, self.uid, self.password,
                        'res.partner', 'read',
                        [batch_partner_ids],  # wrap in a list
                        {'fields': ['name', 'email']}
                    )
                    partners_list.extend(batch_partners)
                    
                    # Update progress
                    if progress_callback:
                        progress = 0.7 + (i / len(partner_ids)) * 0.2
                        progress_callback(f"Fetched {min(i + batch_size, len(partner_ids))}/{len(partner_ids)} clients...", progress)
                
                # Create a dictionary for quick lookup
                partners_data = {partner['id']: partner for partner in partners_list}
            
            # Batch fetch all currencies at once
            currencies_data = {}
            if currency_ids:
                if progress_callback:
                    progress_callback(f"Fetching {len(currency_ids)} currency records...", 0.85)
                
                currencies_list = []
                for i in range(0, len(currency_ids), batch_size):
                    batch_currency_ids = currency_ids[i:i + batch_size]
                    batch_currencies = self.models.execute_kw(
                        self.database, self.uid, self.password,
                        'res.currency', 'read',
                        [batch_currency_ids],  # wrap in a list
                        {'fields': ['name', 'symbol']}
                    )
                    currencies_list.extend(batch_currencies)
                    
                    # Update progress
                    if progress_callback:
                        progress = 0.85 + (i / len(currency_ids)) * 0.1
                        progress_callback(f"Fetched {min(i + batch_size, len(currency_ids))}/{len(currency_ids)} currencies...", progress)
                
                # Create a dictionary for quick lookup
                currencies_data = {currency['id']: currency for currency in currencies_list}
            
            # Batch fetch all companies at once
            companies_data = {}
            if company_ids:
                if progress_callback:
                    progress_callback(f"Fetching {len(company_ids)} company records...", 0.9)
                
                companies_list = []
                for i in range(0, len(company_ids), batch_size):
                    batch_company_ids = company_ids[i:i + batch_size]
                    batch_companies = self.models.execute_kw(
                        self.database, self.uid, self.password,
                        'res.company', 'read',
                        [batch_company_ids],  # wrap in a list
                        {'fields': ['name']}
                    )
                    companies_list.extend(batch_companies)
                    
                    # Update progress
                    if progress_callback:
                        progress = 0.9 + (i / len(company_ids)) * 0.05
                        progress_callback(f"Fetched {min(i + batch_size, len(company_ids))}/{len(company_ids)} companies...", progress)
                
                # Create a dictionary for quick lookup
                companies_data = {company['id']: company for company in companies_list}
            
            # Process invoices with partner data
            if progress_callback:
                progress_callback("Processing invoice data...", 0.95)
            
            invoices = []
            for i, invoice in enumerate(invoices_data):
                if invoice.get('partner_id') and invoice['partner_id'][0] in partners_data:
                    partner = partners_data[invoice['partner_id'][0]]
                    
                    # Calculate days overdue
                    due_date = datetime.strptime(invoice['invoice_date_due'], '%Y-%m-%d').date()
                    days_overdue = (today - due_date).days
                    
                    # Get currency information
                    currency_symbol = "$"  # Default fallback
                    if invoice.get('currency_id') and invoice['currency_id'][0] in currencies_data:
                        currency_symbol = currencies_data[invoice['currency_id'][0]]['symbol'] or "$"
                    
                    # Get company information
                    company_name = "Unknown"  # Default fallback
                    if invoice.get('company_id') and invoice['company_id'][0] in companies_data:
                        company_name = companies_data[invoice['company_id'][0]]['name']
                    
                    invoices.append({
                        'invoice_number': invoice['name'],
                        'due_date': invoice['invoice_date_due'],
                        'days_overdue': days_overdue,
                        'amount_due': invoice['amount_residual'],
                        'currency_symbol': currency_symbol,
                        'origin': invoice.get('invoice_origin', ''),
                        'client_name': partner['name'],
                        'client_email': partner['email'] or '',
                        'invoice_id': invoice['id'],
                        'company_name': company_name
                    })
                
                # Update progress for processing
                if progress_callback and i % 100 == 0:
                    progress = 0.95 + (i / len(invoices_data)) * 0.05
                    progress_callback(f"Processed {i}/{len(invoices_data)} invoices...", progress)
            
            return invoices
        except Exception as e:
            st.error(f"Error fetching invoices: {str(e)}")
            return []

def send_email(sender_email, sender_password, recipient_email, cc_list, subject, body, attachments=None, smtp_server="smtp.gmail.com", smtp_port=587):
    """Send email with optional PDF attachments using SMTP"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg['Cc'] = ', '.join(cc_list) if cc_list else ''
        
        # Add text part
        text_part = MIMEText(body, 'html')  # Changed to HTML to support formatting
        msg.attach(text_part)
        
        # Add attachments if provided
        if attachments:
            for attachment in attachments:
                if attachment is not None:
                    # Create attachment
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    
                    # Set filename
                    filename = attachment.name
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {filename}'
                    )
                    msg.attach(part)
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        recipients = [recipient_email] + cc_list
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending email: {str(e)}")
        return False

def generate_simple_email_template(client_name, invoices, template_type="initial"):
    """Generate simple email template"""
    total_amount = sum(invoice['amount_due'] for invoice in invoices)
    currency_symbol = invoices[0]['currency_symbol'] if invoices else "$"
    max_days = max(inv['days_overdue'] for inv in invoices)
    
    # Simple table
    table_lines = []
    table_lines.append("| Invoice | Due Date | Days Overdue | Amount |")
    table_lines.append("|---------|----------|--------------|--------|")
    
    for inv in invoices:
        table_lines.append(f"| {inv['invoice_number']} | {inv['due_date']} | {inv['days_overdue']} days | {currency_symbol}{inv['amount_due']:,.2f} |")
    
    table_text = '\n'.join(table_lines)
    
    # Simple templates
    templates = {
        "initial": {
            "subject": f"Payment Reminder - {client_name}",
            "body": f"""
            <h2>Dear {client_name},</h2>
            <p>This is a friendly reminder that you have outstanding invoices totaling <strong>{currency_symbol}{total_amount:,.2f}</strong>.</p>
            <p>Please find the details below:</p>
            <pre>{table_text}</pre>
            <p>Please arrange payment at your earliest convenience.</p>
            <p>Best regards,<br>Finance Team</p>
            """
        },
        "second": {
            "subject": f"Urgent Payment Reminder - {client_name}",
            "body": f"""
            <h2>Dear {client_name},</h2>
            <p>This is an urgent reminder that you have outstanding invoices totaling <strong>{currency_symbol}{total_amount:,.2f}</strong> that are overdue by up to {max_days} days.</p>
            <p>Please find the details below:</p>
            <pre>{table_text}</pre>
            <p>Please arrange immediate payment to avoid any further action.</p>
            <p>Best regards,<br>Finance Team</p>
            """
        },
        "final": {
            "subject": f"Final Payment Notice - {client_name}",
            "body": f"""
            <h2>Dear {client_name},</h2>
            <p>This is a final notice regarding your outstanding invoices totaling <strong>{currency_symbol}{total_amount:,.2f}</strong> that are overdue by up to {max_days} days.</p>
            <p>Please find the details below:</p>
            <pre>{table_text}</pre>
            <p>Immediate payment is required to avoid escalation.</p>
            <p>Best regards,<br>Finance Team</p>
            """
        }
    }
    
    return templates[template_type]["subject"], templates[template_type]["body"]

# Sidebar
with st.sidebar:
    st.markdown('<div class="sidebar-header">📧 Email Configuration</div>', unsafe_allow_html=True)
    
    # Email configuration
    default_sender_email = get_secret("EMAIL", "")
    sender_email = st.text_input("Sender Email", value=default_sender_email, help="Email address to send from")
    
    default_sender_password = get_secret("EMAIL_PASSWORD", "")
    sender_password = st.text_input("Sender Password", type="password", value=default_sender_password, help="Password for sender email account")
    if sender_password:
        st.session_state.sender_password = sender_password
    
    st.markdown('<div class="sidebar-header">🔗 Odoo Connection</div>', unsafe_allow_html=True)
    
    # Demo mode option
    demo_mode = st.checkbox("🧪 Demo Mode", help="Use demo data for testing (no Odoo connection required)")
    
    if demo_mode:
        st.info("🧪 Demo mode enabled - using sample data")
        if st.button("🎲 Generate Demo Data", type="secondary"):
            with st.spinner("Generating demo data..."):
                # Import demo data generator
                from demo_data import generate_demo_data
                demo_invoices = generate_demo_data()
                st.session_state.odoo_connected = True
                st.session_state.overdue_invoices = demo_invoices
                st.session_state.clients_missing_email = []
                st.success(f"✅ Demo data loaded! Generated {len(demo_invoices)} sample invoices.")
                st.rerun()
    
    # Odoo connection fields
    default_odoo_url = get_secret("ODOO_URL", "https://prezlab-staging-22061821.dev.odoo.com")
    odoo_url = st.text_input("Odoo URL", value=default_odoo_url, help="e.g., https://your-odoo-instance.com")
    
    default_odoo_database = get_secret("ODOO_DB", "prezlab-staging-22061821")
    odoo_database = st.text_input("Database", value=default_odoo_database, help="Odoo database name")
    
    default_odoo_username = get_secret("ODOO_USERNAME", "omar.elhasan@prezlab.com")
    odoo_username = st.text_input("Username", value=default_odoo_username, help="Odoo username")
    
    default_odoo_password = get_secret("ODOO_PASSWORD", "Omar@@1998")
    odoo_password = st.text_input("Password", value=default_odoo_password, type="password", help="Odoo password")
    
    # Connect button
    if st.button("🔗 Connect to Odoo", type="primary"):
        if odoo_url and odoo_database and odoo_username and odoo_password:
            # Create progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Step 1: Connect to Odoo
                status_text.text("Connecting to Odoo...")
                progress_bar.progress(25)
                
                connector = OdooConnector(odoo_url, odoo_database, odoo_username, odoo_password)
                if not connector.connect():
                    st.error("❌ Failed to connect to Odoo. Please check your credentials.")
                    st.stop()
                
                # Step 2: Fetch overdue invoices
                status_text.text("Fetching overdue invoices...")
                progress_bar.progress(50)
                
                # Add a callback to update progress during data fetching
                def update_progress(message, progress):
                    status_text.text(message)
                    # Ensure progress is always between 0.0 and 1.0
                    safe_progress = min(max(progress, 0.0), 1.0)
                    progress_bar.progress(safe_progress)
                
                invoices = connector.get_overdue_invoices(progress_callback=update_progress)
                
                # Step 3: Process data
                status_text.text("Processing invoice data...")
                progress_bar.progress(75)
                
                # Check for clients missing email
                missing_email_clients = [inv for inv in invoices if not inv['client_email']]
                
                # Step 4: Complete
                status_text.text("Connection successful!")
                progress_bar.progress(100)
                time.sleep(0.5)  # Brief pause to show completion
                
                # Store in session state
                st.session_state.odoo_connected = True
                st.session_state.connector = connector
                st.session_state.overdue_invoices = invoices
                st.session_state.clients_missing_email = missing_email_clients
                
                # Clear progress indicators
                progress_bar.empty()
                status_text.empty()
                
                st.success(f"✅ Successfully connected to Odoo! Found {len(invoices)} overdue invoices.")
                
            except Exception as e:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ Connection failed: {str(e)}")
        else:
            st.error("Please fill in all Odoo connection fields.")
    
    # Display connection status and refresh option
    if st.session_state.odoo_connected:
        if demo_mode:
            st.success("✅ Demo mode active")
        else:
            st.success("✅ Connected to Odoo")
            if st.button("🔄 Refresh", type="secondary", help="Refresh invoice data", use_container_width=True):
                progress_bar = st.progress(0)
                status_text = st.empty()
                try:
                    if 'connector' in st.session_state:
                        def update_refresh_progress(message, progress):
                            status_text.text(message)
                            progress_bar.progress(progress)
                        invoices = st.session_state.connector.get_overdue_invoices(progress_callback=update_refresh_progress)
                        st.session_state.overdue_invoices = invoices
                        st.session_state.clients_missing_email = [inv for inv in invoices if not inv['client_email']]
                        progress_bar.empty()
                        status_text.empty()
                        st.success(f"✅ Refreshed! Found {len(invoices)} overdue invoices.")
                        st.rerun()
                except Exception as e:
                    progress_bar.empty()
                    status_text.empty()
                    st.error(f"❌ Refresh failed: {str(e)}")
    else:
        st.warning("⚠️ Not connected to Odoo")
    
    # Warning for clients missing email
    if st.session_state.clients_missing_email:
        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
        st.warning(f"⚠️ {len(st.session_state.clients_missing_email)} client(s) missing email addresses")
        for client in st.session_state.clients_missing_email[:3]:  # Show first 3
            st.write(f"• {client['client_name']}")
        if len(st.session_state.clients_missing_email) > 3:
            st.write(f"... and {len(st.session_state.clients_missing_email) - 3} more")
        st.markdown('</div>', unsafe_allow_html=True)

# Main content
st.markdown('<div class="main-header">📧 Odoo Invoice Follow-Up Manager</div>', unsafe_allow_html=True)

# Create tabs
tab1, tab2 = st.tabs(["📊 Dashboard", "📧 Bulk Email Sender"])

# Tab 1: Dashboard
with tab1:
    st.markdown("## 📊 Overdue Invoices Dashboard")
    
    if not st.session_state.odoo_connected:
        st.info("Please connect to Odoo using the sidebar to view overdue invoices, or enable Demo Mode for testing.")
    else:
        if st.session_state.overdue_invoices:
            # Create DataFrame with caching
            @st.cache_data
            def create_invoice_dataframe(invoices):
                df = pd.DataFrame(invoices)
                # Add clickable invoice links
                df['Invoice Link'] = df['invoice_number'].apply(
                    lambda x: f"<a href='#' target='_blank'>{x}</a>"
                )
                # Format amount with correct currency symbol
                df['Amount Due'] = df.apply(lambda row: f"{row['currency_symbol']}{row['amount_due']:,.2f}", axis=1)
                # Add Reference Company column (using company_name from invoice data)
                df['Reference Company'] = df['company_name']
                return df
            
            df = create_invoice_dataframe(st.session_state.overdue_invoices)
            
            # Display summary with performance metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Overdue Invoices", len(df))
            with col2:
                # Get the most common currency symbol for display
                currency_symbol = df['currency_symbol'].mode().iloc[0] if not df.empty else "$"
                st.metric("Total Amount Due", f"{currency_symbol}{df['amount_due'].sum():,.2f}")
            with col3:
                st.metric("Average Days Overdue", f"{df['days_overdue'].mean():.1f}")
            with col4:
                st.metric("Clients with Overdue Invoices", df['client_name'].nunique())
            
            # Group by overdue period with caching
            @st.cache_data
            def create_grouped_dataframes(df):
                recent = df[df['days_overdue'] <= 15]
                moderate = df[(df['days_overdue'] > 15) & (df['days_overdue'] <= 30)]
                severe = df[df['days_overdue'] > 30]
                return recent, moderate, severe
            
            recent, moderate, severe = create_grouped_dataframes(df)
            
            # Display grouped data
            st.markdown("### 📅 Overdue by Period")
            
            # 15 days or less
            if not recent.empty:
                st.markdown("#### 🔵 Recent (≤ 15 days)")
                st.dataframe(
                    recent[['invoice_number', 'client_name', 'Reference Company', 'origin', 'due_date', 'days_overdue', 'Amount Due']], 
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "invoice_number": st.column_config.TextColumn("Invoice Number", width="medium"),
                        "client_name": st.column_config.TextColumn("Client Name", width="medium"),
                        "Reference Company": st.column_config.TextColumn("Reference Company", width="medium"),
                        "origin": st.column_config.TextColumn("Origin", width="medium"),
                        "due_date": st.column_config.TextColumn("Due Date", width="small"),
                        "days_overdue": st.column_config.NumberColumn("Days Overdue", width="small"),
                        "Amount Due": st.column_config.TextColumn("Amount Due", width="small")
                    }
                )
            
            # 16-30 days
            if not moderate.empty:
                st.markdown("#### 🟡 Moderate (16-30 days)")
                st.dataframe(
                    moderate[['invoice_number', 'client_name', 'Reference Company', 'origin', 'due_date', 'days_overdue', 'Amount Due']], 
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "invoice_number": st.column_config.TextColumn("Invoice Number", width="medium"),
                        "client_name": st.column_config.TextColumn("Client Name", width="medium"),
                        "Reference Company": st.column_config.TextColumn("Reference Company", width="medium"),
                        "origin": st.column_config.TextColumn("Origin", width="medium"),
                        "due_date": st.column_config.TextColumn("Due Date", width="small"),
                        "days_overdue": st.column_config.NumberColumn("Days Overdue", width="small"),
                        "Amount Due": st.column_config.TextColumn("Amount Due", width="small")
                    }
                )
            
            # 31+ days
            if not severe.empty:
                st.markdown("#### 🔴 Severe (31+ days)")
                st.dataframe(
                    severe[['invoice_number', 'client_name', 'Reference Company', 'origin', 'due_date', 'days_overdue', 'Amount Due']], 
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "invoice_number": st.column_config.TextColumn("Invoice Number", width="medium"),
                        "client_name": st.column_config.TextColumn("Client Name", width="medium"),
                        "Reference Company": st.column_config.TextColumn("Reference Company", width="medium"),
                        "origin": st.column_config.TextColumn("Origin", width="medium"),
                        "due_date": st.column_config.TextColumn("Due Date", width="small"),
                        "days_overdue": st.column_config.NumberColumn("Days Overdue", width="small"),
                        "Amount Due": st.column_config.TextColumn("Amount Due", width="small")
                    }
                )
            
            # All invoices table
            st.markdown("### 📋 All Overdue Invoices")
            st.dataframe(
                df[['invoice_number', 'client_name', 'Reference Company', 'origin', 'due_date', 'days_overdue', 'Amount Due']], 
                use_container_width=True,
                hide_index=True,
                column_config={
                    "invoice_number": st.column_config.TextColumn("Invoice Number", width="medium"),
                    "client_name": st.column_config.TextColumn("Client Name", width="medium"),
                    "Reference Company": st.column_config.TextColumn("Reference Company", width="medium"),
                    "origin": st.column_config.TextColumn("Origin", width="medium"),
                    "due_date": st.column_config.TextColumn("Due Date", width="small"),
                    "days_overdue": st.column_config.NumberColumn("Days Overdue", width="small"),
                    "Amount Due": st.column_config.TextColumn("Amount Due", width="small")
                }
            )
        else:
            st.success("🎉 No overdue invoices found!")

# Tab 2: Bulk Email Sender
with tab2:
    st.markdown("## 📧 Bulk Email Sender")
    
    if not st.session_state.odoo_connected:
        st.info("Please connect to Odoo using the sidebar to send bulk emails, or enable Demo Mode for testing.")
    elif not st.session_state.overdue_invoices:
        st.success("🎉 No overdue invoices to send bulk emails for!")
    else:
        # Group invoices by client for multi-selection
        client_invoices = {}
        for invoice in st.session_state.overdue_invoices:
            client_name = invoice['client_name']
            if client_name not in client_invoices:
                client_invoices[client_name] = []
            client_invoices[client_name].append(invoice)
        
        st.markdown("### 👥 Multi-Client Selection")
        
        # Group clients by overdue period for better organization
        recent_clients = []
        moderate_clients = []
        severe_clients = []
        
        for client_name, invoices in client_invoices.items():
            max_days = max(inv['days_overdue'] for inv in invoices)
            if max_days <= 15:
                recent_clients.append(client_name)
            elif max_days <= 30:
                moderate_clients.append(client_name)
            else:
                severe_clients.append(client_name)
        
        # Create multi-select with grouping
        st.markdown("#### Select Clients to Send Emails To:")
        
        selected_clients = []
        
        # Recent clients section
        if recent_clients:
            st.markdown("**🔵 Recent (≤ 15 days):**")
            recent_selected = st.multiselect(
                "Choose from recent overdue clients:",
                recent_clients,
                key="recent_multiselect",
                help="Select clients with invoices overdue 15 days or less"
            )
            selected_clients.extend(recent_selected)
        
        # Moderate clients section
        if moderate_clients:
            st.markdown("**🟡 Moderate (16-30 days):**")
            moderate_selected = st.multiselect(
                "Choose from moderate overdue clients:",
                moderate_clients,
                key="moderate_multiselect",
                help="Select clients with invoices overdue 16-30 days"
            )
            selected_clients.extend(moderate_selected)
        
        # Severe clients section
        if severe_clients:
            st.markdown("**🔴 Severe (31+ days):**")
            severe_selected = st.multiselect(
                "Choose from severely overdue clients:",
                severe_clients,
                key="severe_multiselect",
                help="Select clients with invoices overdue 31+ days"
            )
            selected_clients.extend(severe_selected)
        
        # Show selection summary
        if selected_clients:
            st.markdown(f"### 📋 Selected Clients ({len(selected_clients)})")
            
            # Create summary table
            summary_data = []
            for client in selected_clients:
                invoices = client_invoices[client]
                total_amount = sum(inv['amount_due'] for inv in invoices)
                max_days = max(inv['days_overdue'] for inv in invoices)
                has_email = any(inv['client_email'] for inv in invoices)
                
                # Get currency symbol from first invoice for this client
                currency_symbol = invoices[0]['currency_symbol'] if invoices else "$"
                
                # Get reference company from first invoice
                reference_company = invoices[0]['company_name'] if invoices else "Unknown"
                
                summary_data.append({
                    'Client': client,
                    'Reference Company': reference_company,
                    'Invoices': len(invoices),
                    'Total Amount': f"{currency_symbol}{total_amount:,.2f}",
                    'Max Days Overdue': max_days,
                    'Has Email': "✅" if has_email else "❌"
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True)
            
            # Check for clients without emails
            clients_without_email = [client for client in selected_clients 
                                   if not any(inv['client_email'] for inv in client_invoices[client])]
            
            if clients_without_email:
                st.warning(f"⚠️ The following clients don't have email addresses: {', '.join(clients_without_email)}")
            
            # Email configuration
            st.markdown("### ⚙️ Email Configuration")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Email template selection
                email_template = st.selectbox(
                    "Email Template:",
                    ["initial", "second", "final"],
                    index=0,
                    help="Select the email template to use for all clients"
                )
            
            with col2:
                # CC configuration
                cc_list = st.text_input(
                    "CC List (comma-separated):",
                    value="",
                    help="Emails to CC for all emails"
                )
            
            # Individual Client Email Configuration
            st.markdown("#### ✍️ Individual Client Email Configuration")
            
            # Create tabs for each selected client
            if selected_clients:
                client_tabs = st.tabs([f"📧 {client}" for client in selected_clients])
                
                # Store client-specific configurations
                client_configs = {}
                
                for i, (client, tab) in enumerate(zip(selected_clients, client_tabs)):
                    with tab:
                        client_invoices_list = client_invoices[client]
                        client_email = client_invoices_list[0]['client_email']
                        max_days = max(inv['days_overdue'] for inv in client_invoices_list)
                        
                        # Client info header
                        st.markdown(f"**Client:** {client}")
                        
                        # Tag-based client email input
                        st.markdown(f"**Email for {client}:**")
                        
                        # Filter out invalid emails from the default list
                        def is_valid_email(email):
                            return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email) and email.lower() not in ['na', 'n/a', 'none', '']
                        
                        # Start with empty list if the original email is invalid
                        default_emails = []
                        if client_email and is_valid_email(client_email):
                            default_emails = [client_email]
                        
                        client_emails = st_tags(
                            label=f"Add email addresses for {client}",
                            text="Press enter to add more",
                            value=default_emails,
                            suggestions=[],
                            maxtags=3,
                            key=f"emails_{i}"
                        )
                        
                        # Validate email addresses
                        invalid_emails = [email for email in client_emails if not is_valid_email(email)]
                        if invalid_emails:
                            st.error(f"Invalid email(s) for {client}: {', '.join(invalid_emails)}")
                            st.info("Please remove invalid emails and add valid ones. Valid emails should contain @ and a domain.")
                        
                        # Use the first valid email as the primary email
                        valid_emails = [email for email in client_emails if is_valid_email(email)]
                        client_email = valid_emails[0] if valid_emails else ""
                        
                        st.markdown(f"**Invoices:** {len(client_invoices_list)} | **Total Amount:** ${sum(inv['amount_due'] for inv in client_invoices_list):,.2f}")
                        st.markdown(f"**Max Days Overdue:** {max_days}")
                        
                        # Email configuration in columns
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Individual template selection for this client
                            client_template = st.selectbox(
                                f"Email Template for {client}:",
                                ["initial", "second", "final"],
                                index=0,
                                key=f"template_{i}",
                                help=f"Select email template for {client}"
                            )
                            
                            # Subject for this specific client
                            _, initial_body_text = generate_simple_email_template(client, client_invoices_list, client_template)
                            initial_subject, _ = generate_simple_email_template(client, client_invoices_list, client_template)
                            
                            client_subject = st.text_input(
                                f"Subject for {client}:",
                                value=initial_subject,
                                key=f"subject_{i}",
                                help="Edit the email subject for this specific client"
                            )
                        
                        with col2:
                            # Client-specific CC field
                            client_cc_input = st.text_input(
                                f"CC for {client} (comma-separated):",
                                placeholder="Enter email addresses to CC for this client",
                                key=f"cc_{i}",
                                help="These emails will be CC'd in addition to the company CC list"
                            )
                        
                        # Email body for this specific client
                        st.markdown(f"**Email Body for {client}:**")
                        
                        # Get the email template for this client
                        _, email_template_body = generate_simple_email_template(client, client_invoices_list, client_template)
                        
                        client_email_body = st.text_area(
                            "Email Body:",
                            value=email_template_body,
                            height=250,
                            key=f"body_{i}",
                            help="Edit the email body for this specific client"
                        )
                        
                        # Invoice Table Preview for this client
                        st.markdown(f"**📋 Invoice Table for {client}:**")
                        # Get currency symbol from first invoice for this client
                        currency_symbol = client_invoices_list[0]['currency_symbol'] if client_invoices_list else "$"
                        invoice_data = []
                        for inv in client_invoices_list:
                            # Safely handle origin field - convert to string and handle None/boolean values
                            origin_value = inv.get('origin', '')
                            if origin_value is None:
                                origin_str = ''
                            elif isinstance(origin_value, bool):
                                origin_str = str(origin_value)
                            else:
                                origin_str = str(origin_value)
                            
                            invoice_data.append({
                                'Invoice Number': inv['invoice_number'],
                                'Reference Company': inv['company_name'],
                                'Origin': origin_str,
                                'Due Date': inv['due_date'],
                                'Days Overdue': f"{inv['days_overdue']} days",
                                'Amount Due': f"{currency_symbol}{inv['amount_due']:,.2f}"
                            })
                        
                        df_invoice = pd.DataFrame(invoice_data)
                        # Apply custom styling for better readability
                        st.dataframe(
                            df_invoice, 
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "Invoice Number": st.column_config.TextColumn("Invoice Number", width="medium"),
                                "Reference Company": st.column_config.TextColumn("Reference Company", width="medium"),
                                "Origin": st.column_config.TextColumn("Origin", width="medium"),
                                "Due Date": st.column_config.TextColumn("Due Date", width="small"),
                                "Days Overdue": st.column_config.TextColumn("Days Overdue", width="small"),
                                "Amount Due": st.column_config.TextColumn("Amount Due", width="small")
                            }
                        )
                        
                        # Store configuration for this client
                        client_configs[client] = {
                            'template': client_template,
                            'subject': client_subject,
                            'body': client_email_body,
                            'cc': client_cc_input,
                            'email': client_email,  # Primary email (first valid email)
                            'all_emails': valid_emails,  # All valid emails for this client
                            'invoices': client_invoices_list
                        }
            else:
                st.info("Please select clients above to configure individual emails.")
                client_configs = {}
            
            # Preview and send section
            st.markdown("### 👀 Preview & Send")
            
            if st.button("🔍 Preview All Emails", type="secondary"):
                if not selected_clients:
                    st.error("Please select at least one client.")
                else:
                    st.markdown("### 📧 Email Previews")
                    
                    for i, client in enumerate(selected_clients):
                        if client not in client_configs:
                            st.error(f"❌ {client}: Configuration not found")
                            continue
                            
                        config = client_configs[client]
                        client_invoices_list = config['invoices']
                        client_email = config['email']
                        
                        if not client_email:
                            st.error(f"❌ {client}: No email address available")
                            continue
                        
                        # Use client-specific configuration
                        _, body_text = generate_simple_email_template(client, client_invoices_list, config['template'])
                        
                        # Create expandable preview for each client
                        currency_symbol = client_invoices_list[0]['currency_symbol'] if client_invoices_list else "$"
                        with st.expander(f"📧 {client} - {len(client_invoices_list)} invoice(s) - {currency_symbol}{sum(inv['amount_due'] for inv in client_invoices_list):,.2f}"):
                            # Show all emails for this client
                            if config.get('all_emails'):
                                st.markdown(f"**To:** {', '.join(config['all_emails'])}")
                            else:
                                st.markdown(f"**To:** {client_email}")
                            st.markdown(f"**Subject:** {config['subject']}")
                            
                            # Email Preview Section
                            st.markdown("### 📧 Email Preview")
                            
                            # Use the processed email body directly (already contains all replacements)
                            preview_text = body_text
                            
                            # Use st.text() for simple, reliable display
                            st.text(preview_text)
            
            # Email sending functionality
            st.markdown("### 📤 Send Emails")
            
            if st.button("📧 Send Bulk Emails", type="primary"):
                if not selected_clients:
                    st.error("Please select at least one client.")
                elif not sender_email:
                    st.error("Please configure sender email in the sidebar.")
                else:
                    # Parse CC list
                    cc_emails = [email.strip() for email in cc_list.split(",") if email.strip()]
                    
                    # Progress tracking
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    successful_sends = 0
                    failed_sends = 0
                    failed_clients = []
                    
                    for i, client in enumerate(selected_clients):
                        # Update progress
                        progress = (i + 1) / len(selected_clients)
                        status_text.text(f"Sending email to {client}... ({i + 1}/{len(selected_clients)})")
                        progress_bar.progress(progress)
                        
                        try:
                            # Get client-specific configuration
                            if client not in client_configs:
                                failed_sends += 1
                                failed_clients.append(f"{client} (no configuration)")
                                continue
                                
                            config = client_configs[client]
                            client_invoices_list = config['invoices']
                            client_email = config['email']
                            
                            if not client_email:
                                failed_sends += 1
                                failed_clients.append(f"{client} (no email)")
                                continue
                            
                            # Use client-specific configuration
                            _, body_text = generate_simple_email_template(client, client_invoices_list, config['template'])
                            
                            # Use the processed email body directly (already contains the table)
                            final_email_body = body_text.replace('\n', '<br>')
                            
                            # Combine company CC with client-specific CC
                            client_cc_emails = []
                            if config['cc']:
                                client_cc_emails = [email.strip() for email in config['cc'].split(",") if email.strip()]
                                # Validate client CC emails
                                def is_valid_email(email):
                                    return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email)
                                client_cc_emails = [email for email in client_cc_emails if is_valid_email(email)]
                            
                            all_cc_emails = cc_emails + client_cc_emails
                            all_cc_emails = list(dict.fromkeys(all_cc_emails))  # Remove duplicates
                            
                            # Send email with proper error handling
                            try:
                                # Get sender password from environment or session state
                                sender_password = os.getenv('SENDER_PASSWORD') or st.session_state.get('sender_password', '')
                                
                                if not sender_password:
                                    st.error("Sender password not configured. Please set SENDER_PASSWORD in environment variables.")
                                    break
                                
                                # Send emails to all addresses for this client
                                client_all_emails = config.get('all_emails', [client_email])
                                emails_sent = 0
                                
                                for email_address in client_all_emails:
                                    if email_address and is_valid_email(email_address):
                                        # Send the email with client-specific configuration
                                        send_email(sender_email, sender_password, email_address, all_cc_emails, config['subject'], final_email_body)
                                        emails_sent += 1
                                
                                if emails_sent > 0:
                                    successful_sends += 1
                                else:
                                    failed_sends += 1
                                    failed_clients.append(f"{client} (no valid emails)")
                                
                            except Exception as email_error:
                                failed_sends += 1
                                failed_clients.append(f"{client} (email error: {str(email_error)})")
                                continue
                            
                            # Small delay to prevent overwhelming the email server
                            time.sleep(0.5)
                            
                        except Exception as e:
                            failed_sends += 1
                            failed_clients.append(f"{client} ({str(e)})")
                    
                    # Final status update
                    progress_bar.progress(1.0)
                    status_text.text("Bulk email operation completed!")
                    
                    # Show results
                    if successful_sends > 0:
                        st.success(f"✅ Successfully sent {successful_sends} email(s)")
                    
                    if failed_sends > 0:
                        st.error(f"❌ Failed to send {failed_sends} email(s)")
                        st.markdown("**Failed clients:**")
                        for failed in failed_clients:
                            st.markdown(f"- {failed}")
                    
                    # Summary
                    st.info(f"📊 **Summary:** {successful_sends} successful, {failed_sends} failed out of {len(selected_clients)} total clients")
        else:
            st.info("👆 Select clients from the sections above to send bulk emails.")

# Footer
st.markdown("---")
st.markdown("*Odoo Invoice Follow-Up Manager - Streamlit Application*")
