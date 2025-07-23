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

# Sidebar
with st.sidebar:
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
        st.info("📧 Email functionality will be restored in the next update.")

# Footer
st.markdown("---")
st.markdown("*Odoo Invoice Follow-Up Manager - Streamlit Application*")
