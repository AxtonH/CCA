#!/usr/bin/env python3
"""
Email Templates for Odoo Invoice Follow-Up Manager
This file contains all email templates used in the application.
"""

def get_initial_reminder_template():
    """Get the initial reminder email template"""
    return {
        "subject": "Payment Reminder – Outstanding Balance Notification",
        "body": """Dear [Company Name],

I hope this message finds you well.

It has come to our attention that you have an outstanding balance of [Amount] [Currency] on your account.

We kindly request that you take the necessary steps to settle this amount within the next [Number of Days] days from the date of this email.

If you have already made the payment after receiving this message, please disregard this notice.

Should you have any questions or require assistance, our CS department is available to support you.

[TABLE]

Thank you for your cooperation.

Sincerely,"""
    }

def get_second_reminder_template():
    """Get the second reminder email template"""
    return {
        "subject": "Second Reminder – Outstanding Balance of [Amount] [Currency]",
        "body": """Dear [Company Name],

I hope you're doing well.

This is a follow-up regarding our previous message concerning your outstanding balance of [Amount] [Currency], which remains unsettled on your account.

As mentioned earlier, we kindly request that you arrange for payment within [Number of Days] days to avoid any potential disruptions to services or further escalation.

If you have already processed the payment, please disregard this reminder. Otherwise, we encourage you to reach out to our [Department Name] if you need any assistance or would like to discuss this matter further.

[TABLE]

We appreciate your prompt attention to this issue."""
    }

def get_final_reminder_template():
    """Get the final reminder email template"""
    return {
        "subject": "Third Follow-Up",
        "body": """Dear [Company Name],

This is our final reminder regarding the outstanding balance of [Amount] [Currency] on your account, which has now been overdue for [Number of Days] days.

Despite our previous communications, we have yet to receive payment or hear from your team regarding this matter.

We kindly urge you to settle the balance immediately to prevent further action, which may include suspension of services or referral to collections, in accordance with our company policy.

If payment has been made, please provide confirmation at your earliest convenience. For any concerns, our [Department Name] remains available to assist.

[TABLE]

Thank you for your immediate attention.

Sincerely,"""
    }

def get_template_by_type(template_type):
    """Get template by type"""
    templates = {
        "initial": get_initial_reminder_template(),
        "second": get_second_reminder_template(),
        "final": get_final_reminder_template()
    }
    return templates.get(template_type, get_initial_reminder_template()) 