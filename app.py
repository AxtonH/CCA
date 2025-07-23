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
if 'auto_login_attempted' not in st.session_state:
    st.session_state.auto_login_attempted = False

def attempt_auto_login():
    """Attempt to automatically login to Odoo using environment variables"""
    if st.session_state.auto_login_attempted:
        return
    
    # Get Odoo credentials from environment variables
    odoo_url = get_secret("ODOO_URL", "")
    odoo_database = get_secret("ODOO_DB", "")
    odoo_username = get_secret("ODOO_USERNAME", "")
    odoo_password = get_secret("ODOO_PASSWORD", "")
    
    # Check if all required credentials are available
    if not all([odoo_url, odoo_database, odoo_username, odoo_password]):
        st.session_state.auto_login_attempted = True
        return
    
    try:
        # Create progress indicator for auto-login
        with st.spinner("🔄 Attempting automatic Odoo login..."):
            # Create connector and attempt connection
            connector = OdooConnector(odoo_url, odoo_database, odoo_username, odoo_password)
            if connector.connect():
                # Fetch overdue invoices
                invoices = connector.get_overdue_invoices()
                
                # Check for clients missing email
                missing_email_clients = [inv for inv in invoices if not inv['client_email']]
                
                # Store in session state
                st.session_state.odoo_connected = True
                st.session_state.connector = connector
                st.session_state.overdue_invoices = invoices
                st.session_state.clients_missing_email = missing_email_clients
                
                st.success(f"✅ Automatic login successful! Found {len(invoices)} overdue invoices.")
            else:
                st.warning("⚠️ Automatic login failed. Please check your credentials in the sidebar.")
    except Exception as e:
        st.warning(f"⚠️ Automatic login failed: {str(e)}")
    finally:
        st.session_state.auto_login_attempted = True

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
            st.write(f"[DEBUG] Odoo returned {len(invoice_ids)} overdue invoice IDs.")
            
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
    
    def get_client_reference_company(self, client_name, progress_callback=None):
        """Get the reference company for a client based on their first invoice"""
        try:
            if not self.models:
                return "Unknown"
            
            if progress_callback:
                progress_callback(f"Finding first invoice for {client_name}...", 0.1)
            
            # First, find the partner ID for this client
            partner_ids = self.models.execute_kw(
                self.database, self.uid, self.password,
                'res.partner', 'search',
                [[('name', '=', client_name)]]
            )
            
            if not partner_ids:
                return "Unknown"
            
            partner_id = partner_ids[0]
            
            if progress_callback:
                progress_callback(f"Searching for first invoice for {client_name}...", 0.3)
            
            # Find the first invoice for this client (ordered by creation date)
            invoice_ids = self.models.execute_kw(
                self.database, self.uid, self.password,
                'account.move', 'search',
                [[('partner_id', '=', partner_id), 
                  ('move_type', '=', 'out_invoice'),
                  ('state', '=', 'posted')]],
                {'order': 'create_date asc', 'limit': 1}
            )
            
            if not invoice_ids:
                return "Unknown"
            
            if progress_callback:
                progress_callback(f"Fetching company information for {client_name}...", 0.7)
            
            # Get the invoice with company information
            invoice_data = self.models.execute_kw(
                self.database, self.uid, self.password,
                'account.move', 'read',
                [invoice_ids],
                {'fields': ['company_id']}
            )
            
            if not invoice_data or not invoice_data[0].get('company_id'):
                return "Unknown"
            
            company_id = invoice_data[0]['company_id'][0]
            
            # Get company name
            company_data = self.models.execute_kw(
                self.database, self.uid, self.password,
                'res.company', 'read',
                [[company_id]],
                {'fields': ['name']}
            )
            
            if progress_callback:
                progress_callback(f"Found reference company for {client_name}", 1.0)
            
            if company_data:
                return company_data[0]['name']
            else:
                return "Unknown"
                
        except Exception as e:
            st.error(f"Error getting reference company for {client_name}: {str(e)}")
            return "Unknown"

class InvoicePDFGenerator:
    def __init__(self, odoo_connector):
        self.connector = odoo_connector
        self.driver = None
    
    def generate_client_invoices_pdf(self, client_name, partner_id, progress_callback=None):
        """Generate PDF with all invoices for a client using API-first approach"""
        try:
            if progress_callback:
                progress_callback(f"Generating PDF for {client_name}...", 0.1)
            
            # Try API method first (more reliable and faster)
            pdf_data = self._generate_pdf_via_api(client_name, partner_id, progress_callback)
            if pdf_data:
                if progress_callback:
                    progress_callback(f"PDF generated successfully via API for {client_name}", 1.0)
                return pdf_data
            
            # Only fall back to browser automation if API completely fails
            if progress_callback:
                progress_callback(f"API methods failed, trying browser automation for {client_name}...", 0.5)
            
            # For now, let's skip browser automation and just return None
            # This will force users to rely on the more reliable API method
            if progress_callback:
                progress_callback(f"Browser automation disabled - API method failed for {client_name}", 1.0)
            
            st.warning(f"⚠️ PDF generation failed for {client_name}. Please check if the client has invoices and try again.")
            return None
            
        except Exception as e:
            st.error(f"Error generating PDF for {client_name}: {str(e)}")
            return None
    
    def _generate_pdf_via_api(self, client_name, partner_id, progress_callback=None):
        """Generate PDF using Odoo v17 API with follow-up report filtering"""
        try:
            if not self.connector.models:
                return None
            
            if progress_callback:
                progress_callback(f"Getting partner ID for {client_name}...", 0.1)
            
            # First, get the partner ID if we don't have it
            if isinstance(partner_id, str):
                # partner_id is actually the client name, so we need to find the partner ID
                partner_ids = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'res.partner', 'search',
                    [[('name', '=', partner_id)]]
                )
                if not partner_ids:
                    st.error(f"❌ No partner found for client: {client_name}")
                    return None
                partner_id = partner_ids[0]
                st.success(f"✅ Found partner ID {partner_id} for client: {client_name}")
            
            if progress_callback:
                progress_callback(f"Getting overdue invoice IDs for {client_name}...", 0.3)
            
            # Get only OVERDUE invoice IDs for this client (follow-up report criteria)
            today = datetime.now().date()
            invoice_ids = self.connector.models.execute_kw(
                self.connector.database, self.connector.uid, self.connector.password,
                'account.move', 'search',
                [[('partner_id', '=', partner_id), 
                  ('move_type', '=', 'out_invoice'),
                  ('state', '=', 'posted'),
                  ('payment_state', '!=', 'paid'),
                  ('invoice_date_due', '<', today.isoformat())]]
            )
            
            if not invoice_ids:
                st.error(f"❌ No overdue invoices found for {client_name}")
                return None
            
            st.success(f"✅ Found {len(invoice_ids)} overdue invoices for {client_name}: {invoice_ids}")
            
            if progress_callback:
                progress_callback(f"Found {len(invoice_ids)} overdue invoices for {client_name}...", 0.5)
            
            # Method 1: Try to get report ID first, then use report action
            try:
                if progress_callback:
                    progress_callback(f"Trying to get report ID for {client_name}...", 0.6)
                
                st.info(f"🔍 Trying Method 1: Report action for {client_name}")
                
                # First, find the report action
                report_ids = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'search',
                    [[('report_name', '=', 'account.report_invoice')]]
                )
                
                if report_ids:
                    # Get the report action
                    report_action = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'read',
                        [report_ids[0]],
                        {'fields': ['id', 'name', 'report_name', 'report_type']}
                    )
                    
                    st.success(f"✅ Found report action: {report_action}")
                    
                    # Try to execute the report action
                    pdf_data = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'run',
                        [report_ids[0], invoice_ids]
                    )
                    
                    if pdf_data:
                        import base64
                        if isinstance(pdf_data, str):
                            try:
                                return base64.b64decode(pdf_data)
                            except:
                                pass
                        elif isinstance(pdf_data, bytes):
                            return pdf_data
                else:
                    st.warning(f"⚠️ No report action found for account.report_invoice")
            except Exception as e:
                st.error(f"❌ Method 1 failed: {str(e)}")
                pass  # Silently fail and try next method
            
            # Method 2: Try alternative report names
            try:
                if progress_callback:
                    progress_callback(f"Trying alternative report names for {client_name}...", 0.7)
                
                # Try different report names that might exist
                report_names = [
                    'account.report_invoice_document',
                    'account.report_invoice_with_payments',
                    'account.report_invoice_simple',
                    'account.report_invoice'
                ]
                
                for report_name in report_names:
                    try:
                        # Search for this report
                        report_ids = self.connector.models.execute_kw(
                            self.connector.database, self.connector.uid, self.connector.password,
                            'ir.actions.report', 'search',
                            [[('report_name', '=', report_name)]]
                        )
                        
                        if report_ids:
                            # Try to execute the report
                            pdf_data = self.connector.models.execute_kw(
                                self.connector.database, self.connector.uid, self.connector.password,
                                'ir.actions.report', 'run',
                                [report_ids[0], invoice_ids]
                            )
                            
                            if pdf_data:
                                import base64
                                if isinstance(pdf_data, str):
                                    try:
                                        return base64.b64decode(pdf_data)
                                    except:
                                        pass
                                elif isinstance(pdf_data, bytes):
                                    return pdf_data
                    except Exception as e:
                        continue  # Silently fail and try next report
            except Exception as e:
                pass  # Silently fail and try next method
            
            # Method 3: Try direct HTTP request with proper Odoo v17 authentication
            try:
                if progress_callback:
                    progress_callback(f"Trying direct HTTP request for {client_name}...", 0.8)
                
                st.info(f"🔍 Trying Method 3: HTTP request for {client_name}")
                
                if self.connector.url and invoice_ids:
                    import requests
                    
                    # Create a session to maintain cookies
                    session = requests.Session()
                    session.headers.update({
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
                    })
                    
                    # First, try to authenticate using Odoo v17 method
                    login_url = f"{self.connector.url}/web/session/authenticate"
                    login_data = {
                        'jsonrpc': '2.0',
                        'method': 'call',
                        'params': {
                            'db': self.connector.database,
                            'login': self.connector.username,
                            'password': self.connector.password
                        }
                    }
                    
                    login_response = session.post(login_url, json=login_data, timeout=30)
                    
                    if login_response.status_code == 200:
                        # Try to parse the login response
                        try:
                            login_result = login_response.json()
                            if login_result.get('result', {}).get('uid'):
                                st.success(f"✅ Login successful for {client_name}")
                            else:
                                st.error(f"❌ Login failed for {client_name}: {login_result}")
                        except:
                            st.error(f"❌ Could not parse login response for {client_name}")
                        
                        # Try the report URL with Odoo v17 format
                        report_url = f"{self.connector.url}/report/pdf/account.report_invoice/{','.join(map(str, invoice_ids))}"
                        response = session.get(report_url, timeout=30)
                        
                        if response.status_code == 200:
                            content_type = response.headers.get('content-type', '')
                            if 'application/pdf' in content_type or response.content.startswith(b'%PDF'):
                                st.success(f"✅ HTTP request successful for {client_name} - PDF size: {len(response.content)} bytes")
                                return response.content
                            else:
                                st.error(f"❌ HTTP request returned non-PDF content for {client_name}: {content_type}")
                                st.error(f"❌ Response length: {len(response.content)} bytes")
                                # Show first 200 characters for debugging
                                try:
                                    st.error(f"❌ Response preview: {response.text[:200]}...")
                                except:
                                    st.error("❌ Could not decode response text")
                        else:
                            st.error(f"❌ HTTP request failed for {client_name} with status: {response.status_code}")
                            try:
                                st.error(f"❌ Response content: {response.text[:200]}...")
                            except:
                                st.error("❌ Could not decode error response")
            except Exception as e:
                pass  # Silently fail and try next method
            
            # Method 4: Try to create a custom report action
            try:
                if progress_callback:
                    progress_callback(f"Trying custom report action for {client_name}...", 0.9)
                
                # Try to create a custom report action for the invoices
                action_data = {
                    'name': f'Invoice Report for {client_name}',
                    'report_name': 'account.report_invoice',
                    'report_type': 'qweb-pdf',
                    'model': 'account.move',
                    'data': invoice_ids
                }
                
                # Try to create and execute the action
                action_id = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'create',
                    [action_data]
                )
                
                if action_id:
                    # Try to execute the custom action
                    pdf_data = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'run',
                        [action_id, invoice_ids]
                    )
                    
                    if pdf_data:
                        import base64
                        if isinstance(pdf_data, str):
                            try:
                                return base64.b64decode(pdf_data)
                            except:
                                pass
                        elif isinstance(pdf_data, bytes):
                            return pdf_data
            except Exception as e:
                pass  # Silently fail and try next method
            
            return None
            
        except Exception as e:
            return None
    
    def _generate_pdf_via_browser(self, client_name, partner_id, progress_callback=None):
        """Generate PDF using browser automation"""
        try:
            if progress_callback:
                progress_callback(f"Starting browser for {client_name}...", 0.3)
            
            # Setup browser
            self._setup_browser()
            
            if progress_callback:
                progress_callback(f"Navigating to Odoo for {client_name}...", 0.4)
            
            # Navigate to Odoo
            self.driver.get(f"{self.connector.url}/web")
            
            # Login
            self._login_to_odoo()
            
            if progress_callback:
                progress_callback(f"Finding invoices for {client_name}...", 0.6)
            
            # Navigate to invoices
            self._navigate_to_invoices()
            
            # Navigate to client and click invoices button
            self._navigate_to_client_invoices(client_name)
            
            if progress_callback:
                progress_callback(f"Selecting invoices for {client_name}...", 0.8)
            
            # Select all invoices
            self._select_all_invoices()
            
            if progress_callback:
                progress_callback(f"Generating PDF for {client_name}...", 0.9)
            
            # Print/Download PDF
            pdf_data = self._download_pdf()
            
            return pdf_data
            
        except Exception as e:
            st.error(f"Browser automation failed for {client_name}: {str(e)}")
            
            # Add debugging information
            try:
                if self.driver:
                    st.error(f"Current URL: {self.driver.current_url}")
                    st.error(f"Page title: {self.driver.title}")
                    
                    # Take screenshot for debugging
                    screenshot_path = os.path.join(tempfile.gettempdir(), f"debug_{client_name.replace(' ', '_')}.png")
                    self.driver.save_screenshot(screenshot_path)
                    st.error(f"Debug screenshot saved to: {screenshot_path}")
            except:
                pass
            
            return None
        finally:
            self._cleanup_browser()
    
    def _setup_browser(self):
        """Setup browser with appropriate options"""
        try:
            # Try Chrome first
            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_experimental_option("prefs", {
                "download.default_directory": tempfile.gettempdir(),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            })
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            return
            
        except Exception as e:
            try:
                # Try Firefox
                firefox_options = webdriver.FirefoxOptions()
                firefox_options.add_argument("--headless")
                
                service = Service(GeckoDriverManager().install())
                self.driver = webdriver.Firefox(service=service, options=firefox_options)
                return
                
            except Exception as e2:
                raise Exception(f"Failed to setup browser: {str(e2)}")
    
    def _login_to_odoo(self):
        """Login to Odoo"""
        try:
            # Wait for login form
            wait = WebDriverWait(self.driver, 10)
            username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
            
            # Enter credentials
            username_field.send_keys(self.connector.username)
            password_field = self.driver.find_element(By.ID, "password")
            password_field.send_keys(self.connector.password)
            
            # Submit
            submit_button = self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
            submit_button.click()
            
            # Wait for login to complete
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "o_main_navbar")))
            
        except Exception as e:
            raise Exception(f"Login failed: {str(e)}")
    
    def _navigate_to_invoices(self):
        """Navigate directly to invoice list"""
        try:
            # Navigate directly to the invoice list with the correct action ID
            self.driver.get(f"{self.connector.url}/web#action=325&model=account.move&view_type=list&menu_id=473")
            time.sleep(3)
            
        except Exception as e:
            raise Exception(f"Navigation failed: {str(e)}")
    
    def _navigate_to_client_invoices(self, client_name):
        """Filter invoices by client name in the invoice list"""
        try:
            wait = WebDriverWait(self.driver, 15)
            
            # Wait for the invoice list to load
            time.sleep(3)
            
            # Try to find search field to filter by client
            search_field = None
            search_selectors = [
                "input[placeholder*='Search']",
                "input[placeholder*='search']",
                "input[type='text']",
                "input[class*='search']",
                ".o_searchview_input",
                "input.o_searchview_input",
                ".o_searchview .o_searchview_input"
            ]
            
            for selector in search_selectors:
                try:
                    search_field = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    break
                except:
                    continue
            
            if search_field:
                # Clear and enter client name
                search_field.clear()
                search_field.send_keys(client_name)
                time.sleep(3)
                
                # Try to find and click on the client filter option
                client_filter = None
                client_selectors = [
                    f"//span[contains(text(), '{client_name}')]",
                    f"//div[contains(text(), '{client_name}')]",
                    f"//a[contains(text(), '{client_name}')]",
                    f"//li[contains(text(), '{client_name}')]"
                ]
                
                for selector in client_selectors:
                    try:
                        client_filter = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        break
                    except:
                        continue
                
                if client_filter:
                    client_filter.click()
                    time.sleep(3)
                    return True
                else:
                    # If we can't find the exact client, try to select the first option
                    try:
                        first_option = self.driver.find_element(By.CSS_SELECTOR, ".o_searchview_autocomplete li")
                        first_option.click()
                        time.sleep(3)
                        return True
                    except:
                        raise Exception(f"Could not find client filter for: {client_name}")
            else:
                raise Exception("Could not find search field")
            
        except Exception as e:
            raise Exception(f"Client filter failed: {str(e)}")
    
    def _select_all_invoices(self):
        """Select all visible invoices"""
        try:
            wait = WebDriverWait(self.driver, 10)
            
            # Wait for the invoice list to load
            time.sleep(3)
            
            # Check if invoices are already selected (look for "X selected" text)
            try:
                selected_text = self.driver.find_element(By.XPATH, "//*[contains(text(), 'selected')]")
                if selected_text and "selected" in selected_text.text:
                    # Invoices are already selected, we can proceed
                    return True
            except:
                pass
            
            # Try multiple ways to find the select all checkbox
            select_all_checkbox = None
            checkbox_selectors = [
                "thead input[type='checkbox']",
                ".o_list_view thead input[type='checkbox']",
                ".o_list_table thead input[type='checkbox']",
                "input[type='checkbox']",
                "input[class*='checkbox']"
            ]
            
            for selector in checkbox_selectors:
                try:
                    checkboxes = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for checkbox in checkboxes:
                        if checkbox.is_displayed() and checkbox.is_enabled():
                            select_all_checkbox = checkbox
                            break
                    if select_all_checkbox:
                        break
                except:
                    continue
            
            if select_all_checkbox:
                select_all_checkbox.click()
                time.sleep(2)
                
                # Verify selection worked
                try:
                    selected_text = self.driver.find_element(By.XPATH, "//*[contains(text(), 'selected')]")
                    if selected_text and "selected" in selected_text.text:
                        return True
                except:
                    pass
            else:
                # If no checkbox found, try to select individual invoices
                try:
                    individual_checkboxes = self.driver.find_elements(By.CSS_SELECTOR, "tbody input[type='checkbox']")
                    for checkbox in individual_checkboxes:
                        if checkbox.is_displayed() and checkbox.is_enabled():
                            checkbox.click()
                            time.sleep(0.5)
                except:
                    pass
            
            return True
            
        except Exception as e:
            raise Exception(f"Selection failed: {str(e)}")
    
    def _download_pdf(self):
        """Download PDF and return data"""
        try:
            wait = WebDriverWait(self.driver, 10)
            
            # Try multiple ways to find the print button
            print_button = None
            print_selectors = [
                "//button[text()='Print']",
                "//button[contains(text(), 'Print')]",
                "//button[contains(text(), 'print')]",
                "//a[contains(text(), 'Print')]",
                "//a[contains(text(), 'print')]",
                "//button[@title*='Print']",
                "//button[@aria-label*='Print']",
                "//button[contains(@class, 'print')]",
                "//button[contains(@class, 'Print')]",
                "//button[contains(@class, 'btn-primary')]"
            ]
            
            for selector in print_selectors:
                try:
                    print_button = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    break
                except:
                    continue
            
            if not print_button:
                # Try to find any button that might be a print button
                try:
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    for button in buttons:
                        button_text = button.text.lower()
                        button_title = button.get_attribute("title") or ""
                        if "print" in button_text or "print" in button_title:
                            print_button = button
                            break
                except:
                    pass
            
            if not print_button:
                # Try to find print link
                try:
                    links = self.driver.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        link_text = link.text.lower()
                        if "print" in link_text:
                            print_button = link
                            break
                except:
                    pass
            
            if print_button:
                print_button.click()
                time.sleep(3)
                
                # Wait for PDF to download
                time.sleep(8)
                
                # Look for downloaded file in temp directory
                temp_dir = tempfile.gettempdir()
                pdf_files = [f for f in os.listdir(temp_dir) if f.endswith('.pdf')]
                
                if pdf_files:
                    # Get the most recent PDF file
                    latest_pdf = max(pdf_files, key=lambda f: os.path.getctime(os.path.join(temp_dir, f)))
                    pdf_path = os.path.join(temp_dir, latest_pdf)
                    
                    # Read PDF data
                    with open(pdf_path, 'rb') as f:
                        pdf_data = f.read()
                    
                    # Clean up
                    os.remove(pdf_path)
                    
                    return pdf_data
                else:
                    # If no PDF found, try to get PDF from browser
                    try:
                        # Check if PDF opened in new tab
                        if len(self.driver.window_handles) > 1:
                            self.driver.switch_to.window(self.driver.window_handles[-1])
                            time.sleep(2)
                            
                            # Try to get PDF data from current page
                            page_source = self.driver.page_source
                            if "application/pdf" in page_source or "pdf" in self.driver.current_url.lower():
                                # PDF might be embedded in the page
                                return None  # We'll need to handle this differently
                    except:
                        pass
            
            return None
            
        except Exception as e:
            raise Exception(f"Download failed: {str(e)}")
    
    def _cleanup_browser(self):
        """Clean up browser resources"""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
    
    def get_available_reports(self):
        """Get list of available reports for debugging in Odoo v17"""
        try:
            if not self.connector.models:
                return []
            
            # Method 1: Try to get reports from ir.actions.report (Odoo v17)
            try:
                report_ids = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'search',
                    [[('model', '=', 'account.move')]]
                )
                
                if report_ids:
                    reports = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'read',
                        [report_ids],
                        {'fields': ['name', 'report_name', 'report_type']}
                    )
                    return reports
            except Exception as e:
                pass
            
            # Method 2: Try to get all reports and filter by model
            try:
                all_report_ids = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'search',
                    [[]]
                )
                
                if all_report_ids:
                    all_reports = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'read',
                        [all_report_ids[:20]],  # Limit to first 20
                        {'fields': ['name', 'report_name', 'report_type', 'model']}
                    )
                    # Filter for invoice-related reports
                    invoice_reports = [r for r in all_reports if r.get('model') == 'account.move' or 'invoice' in (r.get('name', '') + r.get('report_name', '')).lower()]
                    return invoice_reports
            except Exception as e:
                pass
            
            # Method 3: Try to get reports by name pattern
            try:
                report_ids = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'search',
                    [[('report_name', 'ilike', 'invoice')]]
                )
                
                if report_ids:
                    reports = self.connector.models.execute_kw(
                        self.connector.database, self.connector.uid, self.connector.password,
                        'ir.actions.report', 'read',
                        [report_ids],
                        {'fields': ['name', 'report_name', 'report_type', 'model']}
                    )
                    return reports
            except Exception as e:
                pass
            
            return []
            
        except Exception as e:
            st.write(f"[DEBUG] Error getting available reports: {str(e)}")
            return []
    
    def get_available_methods(self):
        """Get list of available methods on ir.actions.report for debugging"""
        try:
            if not self.connector.models:
                return []
            
            # Try to get available methods
            try:
                methods = self.connector.models.execute_kw(
                    self.connector.database, self.connector.uid, self.connector.password,
                    'ir.actions.report', 'fields_get',
                    []
                )
                return methods
            except Exception as e:
                return []
            
        except Exception as e:
            return []

def show_pdf_preview(pdf_data, filename):
    """Show PDF preview in a modal-like display"""
    try:
        # Create a temporary file for the PDF
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_data)
            tmp_file_path = tmp_file.name
        
        # Display PDF using streamlit
        with open(tmp_file_path, "rb") as pdf_file:
            st.download_button(
                label=f"📄 Download {filename}",
                data=pdf_file.read(),
                file_name=filename,
                mime="application/pdf"
            )
        
        # Clean up
        os.unlink(tmp_file_path)
        
    except Exception as e:
        st.error(f"Error showing PDF preview: {str(e)}")

def download_pdf(pdf_data, filename):
    """Provide PDF download functionality"""
    try:
        st.download_button(
            label=f"⬇️ Download {filename}",
            data=pdf_data,
            file_name=filename,
            mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Error downloading PDF: {str(e)}")

def validate_client_data_isolation(client_name, invoices):
    """Validate that all invoices belong to the specified client to prevent data leakage"""
    for invoice in invoices:
        if invoice['client_name'] != client_name:
            raise ValueError(f"Data isolation violation: Invoice {invoice['invoice_number']} belongs to {invoice['client_name']}, not {client_name}")
    return True

def generate_email_template(client_name, invoices, days_overdue, template_type="initial"):
    """Generate email template based on template type using new template system"""
    
    validate_client_data_isolation(client_name, invoices)
    total_amount = sum(invoice['amount_due'] for invoice in invoices)
    
    # Get currency symbol from first invoice (assuming all invoices for a client have same currency)
    currency_symbol = invoices[0]['currency_symbol'] if invoices else "$"
    
    # Calculate days for payment (default 30 days for initial reminder)
    payment_days = 30 if template_type == "initial" else 15 if template_type == "second" else 7

    # Create table in the format shown in the image
    table_lines = []
    table_lines.append("| Reference | Date | Due Date | Origin | Total Due |")
    table_lines.append("|-----------|------|----------|--------|-----------|")
    
    for inv in invoices:
        # Format the data according to the image format
        reference = inv['invoice_number']
        date = inv['due_date']  # Using due date as invoice date for now
        due_date = inv['due_date']
        origin = inv.get('origin', '')[:10] if inv.get('origin') else ''  # Truncate origin
        total_due = f"{currency_symbol}{inv['amount_due']:,.2f}"
        
        table_lines.append(f"| {reference} | {date} | {due_date} | {origin} | {total_due} |")
    
    # Add summary rows
    table_lines.append(f"| | | | **Total Due** | **{currency_symbol}{total_amount:,.2f}** |")
    table_lines.append(f"| | | | **Total Overdue** | **{currency_symbol}{total_amount:,.2f}** |")
    
    table_text = '\n'.join(table_lines)

    # Import and use the new template system
    from email_templates import get_template_by_type
    
    template = get_template_by_type(template_type)
    subject = template["subject"]
    body = template["body"]
    
    # Replace placeholders in the template
    body = body.replace("[Company Name]", client_name)
    body = body.replace("[Amount]", f"{total_amount:,.2f}")
    body = body.replace("[Currency]", currency_symbol)
    body = body.replace("[Number of Days]", str(payment_days))
    body = body.replace("[Department Name]", "CS department")
    body = body.replace("[TABLE]", table_text)
    
    # Replace placeholders in the subject line
    subject = subject.replace("[Amount]", f"{total_amount:,.2f}")
    subject = subject.replace("[Currency]", currency_symbol)
    
    return subject, body

def get_client_classification(max_days_overdue):
    """Determine client classification based on maximum days overdue"""
    if max_days_overdue <= 15:
        return "recent"
    elif max_days_overdue <= 30:
        return "moderate"
    else:
        return "severe"

def get_template_by_classification(classification):
    """Get template type based on client classification"""
    template_mapping = {
        "recent": "initial",
        "moderate": "second", 
        "severe": "final"
    }
    return template_mapping.get(classification, "initial")

def get_automatic_iban_attachment(reference_company):
    """Get automatic IBAN letter attachment based on reference company"""
    import os
    import tempfile
    
    # Define the mapping of companies to their IBAN letter files
    iban_letter_mapping = {
        "Prezlab FZ LLC": "IBAN Letter _ Prezlab FZ LLC .pdf",
        "Prezlab Advanced Design Company": "IBAN Letter _ Prezlab Advanced Design Company .pdf"
    }
    
    # Get the filename for the company
    filename = iban_letter_mapping.get(reference_company)
    if not filename:
        return None
    
    # Construct the full file path
    file_path = os.path.join(os.getcwd(), filename)
    
    # Check if file exists
    if not os.path.exists(file_path):
        st.warning(f"⚠️ IBAN letter file not found for {reference_company}: {file_path}")
        return None
    
    try:
        # Read the file content into memory and create a temporary file-like object
        with open(file_path, 'rb') as file:
            file_content = file.read()
        
        # Create a BytesIO object that can be used multiple times
        import io
        file_obj = io.BytesIO(file_content)
        file_obj.name = filename  # Set the filename for the email attachment
        
        return file_obj
    except Exception as e:
        st.error(f"Error reading IBAN letter file for {reference_company}: {str(e)}")
        return None

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

# Attempt automatic login on app launch (after all classes are defined)
attempt_auto_login()

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
    cc_list = st.text_input("CC List", value="", help="Comma-separated email addresses")
    
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
                
                # Check if automatic IBAN attachment will be included
                automatic_iban = "🏦" if reference_company in ["Prezlab FZ LLC", "Prezlab Advanced Design Company"] else ""
                
                # Check if PDF attachment is enabled (will be determined during preview)
                pdf_status = "📄" if ('enable_all_pdfs' in locals() and enable_all_pdfs) or ('enable_all_pdfs' not in locals()) else ""
                
                summary_data.append({
                    'Client': client,
                    'Reference Company': reference_company,
                    'Invoices': len(invoices),
                    'Total Amount': f"{currency_symbol}{total_amount:,.2f}",
                    'Max Days Overdue': max_days,
                    'Has Email': "✅" if has_email else "❌",
                    'Auto IBAN': automatic_iban,
                    'Auto PDF': pdf_status
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True)
            
            # Check for clients without emails
            clients_without_email = [client for client in selected_clients 
                                   if not any(inv['client_email'] for inv in client_invoices[client])]
            
            if clients_without_email:
                st.warning(f"⚠️ The following clients don't have email addresses: {', '.join(clients_without_email)}")
                st.info("You can add email addresses manually in the individual email drafter tab.")
            
            # Automatic IBAN Attachment Information
            st.markdown("### 🏦 Automatic IBAN Letter Attachments")
            st.info("""
            **Automatic IBAN Letter Attachments:**
            - **Prezlab FZ LLC** clients will automatically receive: `IBAN Letter _ Prezlab FZ LLC .pdf`
            - **Prezlab Advanced Design Company** clients will automatically receive: `IBAN Letter _ Prezlab Advanced Design Company .pdf`
            - These attachments are added automatically based on the client's Reference Company
            - Manual and client-specific attachments are still supported
            """)
            
            # Automatic Invoice PDF Attachment Information
            st.markdown("### 📄 Automatic Invoice PDF Attachments")
            st.info("""
            **Automatic Invoice PDF Attachments:**
            - Generate PDF containing all invoices for each client
            - PDFs are generated during email preview and can be previewed/downloaded
            - Each client receives their own confidential PDF with all their invoices
            - Enable/disable per client using the toggles below
            - If PDF generation fails, you'll be prompted to attach invoices manually
            """)
            
            # Master toggle for all clients
            enable_all_pdfs = st.checkbox(
                "📄 Enable automatic invoice PDF attachment for all clients",
                value=True,
                help="Master toggle to enable/disable PDF generation for all selected clients"
            )
            
            # Bulk Email Template Configuration
            st.markdown("### ⚙️ Bulk Email Template Configuration")
            
            # Template selection and CC configuration in the same row
            col1, col2 = st.columns(2)
            
            with col1:
                # Bulk template selection that applies to all clients
                bulk_template = st.selectbox(
                    "Email Template for All Clients:",
                    ["initial", "second", "final"],
                    index=0,  # Default to initial reminder
                    help="This template will be applied to all selected clients"
                )
            
            with col2:
                cc_list_bulk = st.text_input(
                    "CC List (comma-separated):",
                    value=cc_list if 'cc_list' in locals() else "",
                    help="Emails to CC for all bulk emails"
                )
                
                # Show sender info
                if sender_email:
                    st.info(f"📧 Emails will be sent from: {sender_email}")
                else:
                    st.warning("⚠️ No sender email configured")
            
            # PDF Attachments Configuration
            st.markdown("#### 📎 PDF Attachments")
            uploaded_files = st.file_uploader(
                "Upload PDF files to attach to all emails:",
                type=['pdf'],
                accept_multiple_files=True,
                help="Select PDF files to attach to all emails. These will be sent to all selected clients."
            )
            
            if uploaded_files:
                st.success(f"✅ {len(uploaded_files)} PDF file(s) selected for attachment:")
                for file in uploaded_files:
                    st.write(f"• {file.name} ({file.size} bytes)")
            
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
                        
                        # PDF attachment toggle for this client
                        enable_pdf_attachment = st.checkbox(
                            f"📄 Enable automatic invoice PDF for {client}",
                            value=enable_all_pdfs if 'enable_all_pdfs' in locals() else True,
                            key=f"pdf_toggle_{i}",
                            help=f"Generate PDF with all invoices for {client}"
                        )
                        
                        # Email configuration in columns
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Determine client classification and auto-set template
                            client_classification = get_client_classification(max_days)
                            auto_template = get_template_by_classification(client_classification)
                            
                            # Show classification info
                            classification_emoji = {"recent": "🔵", "moderate": "🟡", "severe": "🔴"}
                            st.info(f"{classification_emoji.get(client_classification, '⚪')} **Classification:** {client_classification.title()} ({max_days} days overdue)")
                            
                            # Individual template selection for this client (auto-selected based on classification)
                            client_template = st.selectbox(
                                f"Email Template for {client}:",
                                ["initial", "second", "final"],
                                index=["initial", "second", "final"].index(auto_template),
                                key=f"template_{i}",
                                help=f"Auto-selected '{auto_template}' template based on {client_classification} classification. You can override if needed."
                            )
                            
                            # Subject for this specific client (using individual template)
                            _, initial_body_text = generate_email_template(client, client_invoices_list, max_days, client_template)
                            initial_subject, _ = generate_email_template(client, client_invoices_list, max_days, client_template)
                            
                            # Make client name bold in subject
                            initial_subject = initial_subject.replace(client, f"**{client}**")
                            
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
                        
                        # Create email template for this client using the new template system
                        from email_templates import get_template_by_type
                        template = get_template_by_type(client_template)
                        # Keep all placeholders in the template for editing
                        email_template = template["body"]
                        
                        client_email_body = st.text_area(
                            "Email Body:",
                            value=email_template,
                            height=250,
                            key=f"body_{i}",
                            help="Edit the email body for this specific client. The invoice table will be automatically inserted."
                        )
                        
                        # Individual PDF attachments for this client
                        st.markdown(f"**📎 PDF Attachments for {client}:**")
                        client_attachments = st.file_uploader(
                            f"Upload PDF files for {client}:",
                            type=['pdf'],
                            accept_multiple_files=True,
                            help=f"Select PDF files to attach specifically to {client}'s email"
                        )
                        
                        if client_attachments:
                            st.success(f"✅ {len(client_attachments)} PDF file(s) selected for {client}:")
                            for file in client_attachments:
                                st.write(f"• {file.name} ({file.size} bytes)")
                        
                        # Invoice Table Preview for this client
                        st.markdown(f"**📋 Invoice Table for {client}:**")
                        # Get currency symbol from first invoice for this client
                        currency_symbol = client_invoices_list[0]['currency_symbol'] if client_invoices_list else "$"
                        invoice_data = [
                            {
                                'Invoice Number': inv['invoice_number'],
                                'Reference Company': inv['company_name'],
                                'Origin': str(inv.get('origin', '')) if inv.get('origin') else '',
                                'Due Date': inv['due_date'],
                                'Days Overdue': f"{inv['days_overdue']} days",
                                'Amount Due': f"{currency_symbol}{inv['amount_due']:,.2f}"
                            }
                            for inv in client_invoices_list
                        ]
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
                            'template': client_template,  # Use individual template
                            'subject': client_subject,
                            'body': client_email_body,
                            'cc': client_cc_input,
                            'email': client_email,  # Primary email (first valid email)
                            'all_emails': valid_emails,  # All valid emails for this client
                            'invoices': client_invoices_list,
                            'attachments': client_attachments if client_attachments else [],
                            'enable_pdf_attachment': enable_pdf_attachment,  # PDF attachment toggle
                            'generated_pdf': None,  # Store generated PDF data
                            'pdf_generation_status': 'pending'  # pending/success/failed
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
                    
                    # Initialize PDF generator if we have a connector
                    pdf_generator = None
                    if 'connector' in st.session_state:
                        pdf_generator = InvoicePDFGenerator(st.session_state.connector)
                    
                    # PDFs will be generated during email sending for better reliability
                    
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
                        max_days = max(inv['days_overdue'] for inv in client_invoices_list)
                        _, body_text = generate_email_template(client, client_invoices_list, max_days, config['template'])
                        
                        # Create expandable preview for each client
                        currency_symbol = client_invoices_list[0]['currency_symbol'] if client_invoices_list else "$"
                        with st.expander(f"📧 {client} - {len(client_invoices_list)} invoice(s) - {currency_symbol}{sum(inv['amount_due'] for inv in client_invoices_list):,.2f}"):
                            # Show all emails for this client
                            if config.get('all_emails'):
                                st.markdown(f"**To:** {', '.join(config['all_emails'])}")
                            else:
                                st.markdown(f"**To:** {client_email}")
                            st.markdown(f"**Subject:** {config['subject']}")
                            
                            # Show attachments that will be included
                            st.markdown("### 📎 Attachments")
                            attachment_list = []
                            
                            # Manual uploads
                            if uploaded_files:
                                for file in uploaded_files:
                                    attachment_list.append(f"📄 {file.name} (manual upload)")
                            
                            # Client-specific attachments
                            if config['attachments']:
                                for file in config['attachments']:
                                    attachment_list.append(f"📄 {file.name} (client-specific)")
                            
                            # Automatic IBAN letter
                            reference_company = client_invoices_list[0]['company_name'] if client_invoices_list else "Unknown"
                            automatic_iban_attachment = get_automatic_iban_attachment(reference_company)
                            if automatic_iban_attachment:
                                iban_filename = "IBAN Letter _ Prezlab FZ LLC .pdf" if reference_company == "Prezlab FZ LLC" else "IBAN Letter _ Prezlab Advanced Design Company .pdf"
                                attachment_list.append(f"🏦 {iban_filename} (automatic - {reference_company})")
                            
                            # Generated invoice PDF
                            if config.get('enable_pdf_attachment', False):
                                attachment_list.append(f"📄 {client} - invoice.pdf (will be generated during sending)")
                            
                            if attachment_list:
                                for attachment in attachment_list:
                                    st.markdown(f"• {attachment}")
                            else:
                                st.markdown("• No attachments")
                            
                            # Email Preview Section
                            st.markdown("### 📧 Email Preview")
                            
                            # Use the processed email body directly (already contains all replacements)
                            preview_text = body_text
                            
                            # Use st.text() for simple, reliable display
                            st.text(preview_text)
            
            # Send bulk emails
            st.markdown("### 📤 Send Bulk Emails")
            
            if st.button("📧 Send Bulk Emails", type="primary"):
                if not selected_clients:
                    st.error("Please select at least one client.")
                elif not sender_email:
                    st.error("Please configure sender email in the sidebar.")
                else:
                    # Parse CC list
                    cc_emails = [email.strip() for email in cc_list_bulk.split(",") if email.strip()]
                    
                    # Initialize PDF generator for email sending
                    pdf_generator = None
                    if st.session_state.connector:
                        pdf_generator = InvoicePDFGenerator(st.session_state.connector)
                    
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
                            max_days = max(inv['days_overdue'] for inv in client_invoices_list)
                            _, body_text = generate_email_template(client, client_invoices_list, max_days, config['template'])
                            
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
                            
                            # Prepare attachments for all emails
                            all_attachments = []
                            
                            # Add manually uploaded files
                            if uploaded_files:
                                # Reset file pointers for each email
                                for file in uploaded_files:
                                    file.seek(0)  # Reset file pointer to beginning
                                    all_attachments.append(file)
                            
                            # Add client-specific attachments
                            if config['attachments']:
                                all_attachments.extend(config['attachments'])
                            
                            # Add automatic IBAN letter attachment based on reference company
                            reference_company = client_invoices_list[0]['company_name'] if client_invoices_list else "Unknown"
                            automatic_iban_attachment = get_automatic_iban_attachment(reference_company)
                            if automatic_iban_attachment:
                                all_attachments.append(automatic_iban_attachment)
                                st.info(f"📎 Automatically attaching IBAN letter for {reference_company}")
                            
                            # Generate and attach invoice PDF if enabled
                            if config.get('enable_pdf_attachment', False):
                                if not pdf_generator:
                                    st.error(f"❌ PDF generator not available for {client} - Odoo connection required")
                                    continue
                                    
                                st.info(f"🔍 Generating PDF for {client} during email sending...")
                                
                                try:
                                    # Get partner ID for this client
                                    client_invoices_list = config['invoices']
                                    if not client_invoices_list:
                                        st.warning(f"⚠️ No invoices found for {client}")
                                        continue
                                    
                                    # Get partner ID from first invoice
                                    partner_id = client_invoices_list[0].get('partner_id', None)
                                    if not partner_id:
                                        # Try to get partner ID from Odoo
                                        partner_ids = st.session_state.connector.models.execute_kw(
                                            st.session_state.connector.database, 
                                            st.session_state.connector.uid, 
                                            st.session_state.connector.password,
                                            'res.partner', 'search',
                                            [[('name', '=', client)]]
                                        )
                                        if partner_ids:
                                            partner_id = partner_ids[0]
                                        else:
                                            st.error(f"❌ Could not find partner ID for {client}")
                                            continue
                                    
                                    # Generate PDF fresh during email sending
                                    def update_pdf_progress(message, progress):
                                        st.info(f"{client}: {message}")
                                    
                                    pdf_data = pdf_generator.generate_client_invoices_pdf(
                                        client, partner_id, update_pdf_progress
                                    )
                                    
                                    if pdf_data:
                                        # Create BytesIO object for the PDF
                                        pdf_obj = io.BytesIO(pdf_data)
                                        pdf_obj.name = f"{client} - invoice.pdf"
                                        
                                        all_attachments.append(pdf_obj)
                                        st.success(f"✅ PDF generated and attached for {client} ({len(pdf_data)} bytes)")
                                    else:
                                        st.error(f"❌ PDF generation failed for {client} - no data returned")
                                        
                                except Exception as e:
                                    st.error(f"❌ Error generating PDF for {client}: {str(e)}")
                            
                            # Create fresh copies of attachments for this email to avoid file handle issues
                            email_attachments = []
                            for attachment in all_attachments:
                                if hasattr(attachment, 'getvalue'):  # BytesIO object
                                    # Create a fresh copy of the BytesIO object
                                    import io
                                    fresh_attachment = io.BytesIO(attachment.getvalue())
                                    fresh_attachment.name = attachment.name
                                    email_attachments.append(fresh_attachment)
                                elif hasattr(attachment, 'seek'):  # File-like object
                                    # Reset file pointer and create a copy
                                    attachment.seek(0)
                                    email_attachments.append(attachment)
                                else:
                                    email_attachments.append(attachment)
                            
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
                                        send_email(sender_email, sender_password, email_address, all_cc_emails, config['subject'], final_email_body, attachments=email_attachments)
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