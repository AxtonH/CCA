#!/usr/bin/env python3
"""
Demo data generator for testing the Odoo Invoice Follow-Up Manager
This script generates sample overdue invoice data for testing purposes.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import random

def generate_demo_data():
    """Generate sample overdue invoice data for testing"""
    
    # Sample client names
    clients = [
        "Acme Corporation",
        "TechStart Solutions",
        "Global Industries Ltd",
        "Innovation Systems",
        "Digital Dynamics",
        "Future Technologies",
        "Smart Solutions Inc",
        "NextGen Enterprises",
        "Cloud Computing Corp",
        "Data Analytics Pro"
    ]
    
    # Sample client emails
    client_emails = [
        "accounts@acme.com",
        "finance@techstart.com",
        "billing@globalind.com",
        "payments@innovationsys.com",
        "accounts@digitaldynamics.com",
        "finance@futuretech.com",
        "billing@smartsolutions.com",
        "payments@nextgen.com",
        "accounts@cloudcomp.com",
        "finance@dataanalytics.com"
    ]
    
    # Sample companies (including the specific ones for IBAN testing)
    companies = [
        "Prezlab FZ LLC",
        "Prezlab Advanced Design Company",
        "PrezLab",
        "TechCorp",
        "Innovation Inc",
        "Digital Solutions",
        "Future Systems"
    ]
    
    # Generate sample invoices
    invoices = []
    today = datetime.now().date()
    
    for i in range(25):  # Generate 25 sample invoices
        # Random client
        client_idx = random.randint(0, len(clients) - 1)
        client_name = clients[client_idx]
        client_email = client_emails[client_idx]
        
        # Random due date (past dates for overdue invoices)
        days_overdue = random.randint(1, 60)
        due_date = today - timedelta(days=days_overdue)
        
        # Random amount
        amount = round(random.uniform(500, 10000), 2)
        
        # Generate invoice number
        invoice_number = f"INV-{2024}-{str(i+1).zfill(4)}"
        
        # Generate sample origin
        origins = [
            "S00538 TMHB-T55",
            "S01908, S01897", 
            "S01806, S01713",
            "S01564 MOC22005",
            "S01395",
            "S00132",
            "S01952",
            "S01958",
            "S01836",
            "S0150"
        ]
        origin = random.choice(origins)
        
        # Sample currency symbol
        currency_symbol = "SAR"
        
        # Random company
        company_name = random.choice(companies)
        
        invoices.append({
            'invoice_number': invoice_number,
            'due_date': due_date.strftime('%Y-%m-%d'),
            'days_overdue': days_overdue,
            'amount_due': amount,
            'currency_symbol': currency_symbol,
            'origin': origin,
            'client_name': client_name,
            'client_email': client_email,
            'invoice_id': i + 1,
            'company_name': company_name
        })
    
    return invoices

def main():
    """Main function to run the demo data generator"""
    
    st.set_page_config(
        page_title="Demo Data Generator",
        page_icon="🧪",
        layout="wide"
    )
    
    st.markdown("""
    # 🧪 Demo Data Generator
    
    This tool generates sample overdue invoice data for testing the Odoo Invoice Follow-Up Manager.
    
    **Use this when you don't have access to a real Odoo instance for testing.**
    """)
    
    if st.button("🎲 Generate Demo Data", type="primary"):
        with st.spinner("Generating demo data..."):
            demo_invoices = generate_demo_data()
            
            # Store in session state
            st.session_state.odoo_connected = True
            st.session_state.overdue_invoices = demo_invoices
            st.session_state.clients_missing_email = []
            
            st.success("✅ Demo data generated successfully!")
            
            # Display summary
            df = pd.DataFrame(demo_invoices)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Overdue Invoices", len(df))
            with col2:
                st.metric("Total Amount Due", f"${df['amount_due'].sum():,.2f}")
            with col3:
                st.metric("Average Days Overdue", f"{df['days_overdue'].mean():.1f}")
            with col4:
                st.metric("Clients with Overdue Invoices", df['client_name'].nunique())
            
            # Show sample data
            st.markdown("### 📋 Sample Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            st.info("""
            **Next Steps:**
            1. Go back to the main application
            2. The demo data will be available in both Dashboard and Email Drafter tabs
            3. You can test all features with this sample data
            """)
    
    # Show current demo data if available
    if 'overdue_invoices' in st.session_state and st.session_state.overdue_invoices:
        st.markdown("### 📊 Current Demo Data")
        
        df = pd.DataFrame(st.session_state.overdue_invoices)
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Overdue Invoices", len(df))
        with col2:
            st.metric("Total Amount Due", f"${df['amount_due'].sum():,.2f}")
        with col3:
            st.metric("Average Days Overdue", f"{df['days_overdue'].mean():.1f}")
        with col4:
            st.metric("Clients with Overdue Invoices", df['client_name'].nunique())
        
        # Group by overdue period
        st.markdown("### 📅 Overdue by Period")
        
        recent = df[df['days_overdue'] <= 15]
        moderate = df[(df['days_overdue'] > 15) & (df['days_overdue'] <= 30)]
        severe = df[df['days_overdue'] > 30]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("#### 🔵 Recent (≤ 15 days)")
            st.metric("Count", len(recent))
            if not recent.empty:
                st.dataframe(recent[['invoice_number', 'client_name', 'company_name', 'origin', 'days_overdue', 'amount_due']], 
                           use_container_width=True)
        
        with col2:
            st.markdown("#### 🟡 Moderate (16-30 days)")
            st.metric("Count", len(moderate))
            if not moderate.empty:
                st.dataframe(moderate[['invoice_number', 'client_name', 'company_name', 'origin', 'days_overdue', 'amount_due']], 
                           use_container_width=True)
        
        with col3:
            st.markdown("#### 🔴 Severe (31+ days)")
            st.metric("Count", len(severe))
            if not severe.empty:
                st.dataframe(severe[['invoice_number', 'client_name', 'company_name', 'origin', 'days_overdue', 'amount_due']], 
                           use_container_width=True)
        
        if st.button("🗑️ Clear Demo Data"):
            st.session_state.odoo_connected = False
            st.session_state.overdue_invoices = []
            st.session_state.clients_missing_email = []
            st.success("✅ Demo data cleared!")
            st.rerun()
    
    st.markdown("---")
    st.markdown("*Demo Data Generator - For testing purposes only*")

if __name__ == "__main__":
    main() 