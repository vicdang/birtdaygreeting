#!/usr/bin/env python3
"""Send birthday card via Outlook with embedded image."""

import logging
import json
import base64
from pathlib import Path
from birthday_bot import config, mailer
from birthday_bot.roster import load_roster

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuration
CARD_OUTPUT_DIR = Path('out/birthday_cards_today')
EMAIL_CONFIG_FILE = Path('email_config.json')

def load_email_config() -> dict:
    """Load email configuration from email_config.json."""
    if not EMAIL_CONFIG_FILE.exists():
        logger.error(f"Email config file not found: {EMAIL_CONFIG_FILE}")
        raise FileNotFoundError(f"{EMAIL_CONFIG_FILE} not found")
    
    with open(EMAIL_CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def image_to_base64(image_path: str) -> str:
    """Convert image file to base64 string."""
    with open(image_path, 'rb') as f:
        return base64.b64encode(f.read()).decode('utf-8')

def get_birthday_people_emails(card_filenames: list) -> list:
    """
    Extract birthday people's emails from card filenames.
    Card format: card_<member_id>.png
    """
    try:
        members = load_roster('data/members.xlsx')
        member_map = {str(m['member_id']): m for m in members}
        
        cc_emails = []
        for filename in card_filenames:
            # Extract member ID from filename (card_<ID>.png)
            member_id = filename.stem.replace('card_', '')
            
            if member_id in member_map:
                member = member_map[member_id]
                # Try to get email, fallback to other contact info
                email = member.get('email')
                if email and '@' in email:
                    cc_emails.append(email)
                    logger.info(f"  CC'd: {member['full_name']} ({email})")
                else:
                    logger.warning(f"  No valid email for {member['full_name']}")
        
        return cc_emails
    except Exception as e:
        logger.warning(f"Could not extract birthday people emails: {e}")
        return []

def main():
    """Send birthday card via Outlook with embedded image."""
    
    # Load configurations
    config.load_env()
    email_cfg = load_email_config()
    
    # Extract settings
    email = email_cfg['email']
    template = email_cfg['template']
    styling = email_cfg['styling']
    
    # Check if email sending is enabled
    if not email.get('enabled', True):
        logger.info("⏸️  Email sending is DISABLED in config. Set 'enabled: true' to send emails.")
        return False
    
    # Log sender configuration
    sender_email = email.get('sender_email', 'default@tma.com.vn')
    use_alt = email.get('use_alt_sender', False)
    if use_alt:
        alt_email = email.get('alt_sender_email', sender_email)
        logger.info(f"📧 Sender: {email.get('sender_name', 'Team')} <{alt_email}> (alternative)")
    else:
        logger.info(f"📧 Sender: {email.get('sender_name', 'Team')} <{sender_email}>")
    
    # Handle both single string and list of recipients
    recipients = email['recipient']
    if isinstance(recipients, str):
        recipients = [recipients]
    
    subject = email['subject']
    
    # Find ALL cards for today's birthdays
    cards = sorted(list(CARD_OUTPUT_DIR.glob('*.png')))
    
    if not cards:
        logger.error("No birthday cards found in out/birthday_cards_today/")
        return False
    
    logger.info(f"Found {len(cards)} birthday card(s) for today")
    
    try:
        # Convert all images to base64
        logger.info("Embedding card images in email body...")
        cards_base64 = []
        for card_path in cards:
            logger.info(f"  - {card_path.name}")
            card_base64 = image_to_base64(str(card_path))
            cards_base64.append({
                'name': card_path.stem,
                'base64': card_base64
            })
        
        # Build HTML with all cards - one per line with explicit line breaks
        cards_html = ""
        for idx, card in enumerate(cards_base64):
            cards_html += f'<img src="data:image/png;base64,{card["base64"]}" alt="Birthday Card" style="max-width: 100%; height: auto; border-radius: {styling["card_border_radius"]}; box-shadow: {styling["card_shadow"]}; display: block; margin: 0 auto;" />'
            cards_html += '<br><br>'

        
        # Prepare email content with all embedded images
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; background-color: {styling['background_color']}; }}
                .container {{ max-width: 600px; margin: 0 auto; background-color: {styling['container_background']}; padding: 20px; border-radius: 10px; }}
                .header {{ color: {styling['header_color']}; font-size: {styling['header_size']}; font-weight: bold; margin-bottom: 20px; text-align: center; }}
                .content {{ color: {styling['body_color']}; line-height: 1.6; margin-bottom: 20px; text-align: center; }}
                .card-container {{ margin: 20px 0; text-align: center; }}
                .footer {{ color: {styling['footer_color']}; font-size: 12px; text-align: center; margin-top: 20px; border-top: 1px solid {styling['border_color']}; padding-top: 10px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">{template['header_emoji']} {template['header_text']} {template['header_emoji']}</div>
                <div class="content">
                    <p>{template['greeting']}</p>
                    <p>{template['greeting_secondary']}</p>
                </div>
                <div class="card-container">
                    {cards_html}
                </div>
                </br>
                <div class="content">
                    <p>{template['footer_text']}</p>
                    <p style="margin-top: 20px;">{template['closing']}<br>{template['sender']}</p>
                </div>
                <div class="footer">
                    {template['system_footer']}
                </div>
            </div>
        </body>
        </html>
        """
        
        # Build CC list from config
        cc_list = []
        
        # Add birthday people to CC if configured
        if email.get('cc_birthday_people', False):
            logger.info("CC'ing birthday people...")
            birthday_emails = get_birthday_people_emails(cards)
            cc_list.extend(birthday_emails)
        
        # Add additional CC addresses from config
        additional_cc = email.get('additional_cc', [])
        if additional_cc:
            logger.info(f"Adding additional CC recipients: {additional_cc}")
            cc_list.extend(additional_cc)
        
        # Remove duplicates while preserving order
        cc_list = list(dict.fromkeys(cc_list))
        
        # Determine which sender email to use
        sender_email = email.get('sender_email', 'default@tma.com.vn')
        use_alt = email.get('use_alt_sender', False)
        if use_alt:
            sender_email = email.get('alt_sender_email', sender_email)
        
        # Send to all recipients
        logger.info(f"Sending email to {len(recipients)} recipient(s)...")
        if cc_list:
            logger.info(f"With CC to: {', '.join(cc_list)}")
        all_sent = True
        
        for recipient_email in recipients:
            try:
                logger.info(f"Sending to: {recipient_email}...")
                
                success = mailer.send_email_via_outlook(
                    to_address=recipient_email,
                    subject=subject,
                    html_body=html_body,
                    attachments=None,  # No attachments, image is embedded
                    cc_addresses=cc_list if cc_list else None,
                    from_email=sender_email
                )
                
                if success:
                    logger.info(f"✅ Email sent to {recipient_email}")
                else:
                    logger.error(f"❌ Failed to send to {recipient_email}")
                    all_sent = False
                    
            except mailer.MailerError as e:
                logger.error(f"❌ Error sending to {recipient_email}: {e}")
                all_sent = False
        
        if all_sent:
            logger.info(f"✅ All {len(recipients)} emails sent successfully!")
            return True
        else:
            logger.error("❌ Some emails failed to send")
            return False
        
    except (FileNotFoundError, KeyError, json.JSONDecodeError) as e:
        logger.error(f"❌ Configuration error: {e}")
        return False

if __name__ == '__main__':
    success = main()
    exit(0 if success else 1)
