#!/usr/bin/env python3
"""
Test workflow: Generate birthday cards and send emails for tomorrow's birthdays
"""

import logging
import json
from pathlib import Path
from datetime import datetime, timedelta
from birthday_bot.roster import load_roster
from birthday_bot import config
from birthday_bot.renderer import render_card
from birthday_bot.mailer import send_email_via_outlook

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('birthday_bot_tomorrow_test.log'),
        logging.StreamHandler()
    ]
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

def get_birthday_people_emails(card_filenames: list, members: list) -> list:
    """
    Extract birthday people's emails from card filenames.
    Card format: card_<member_id>.png
    """
    try:
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
    """Generate and send birthday cards for tomorrow's birthdays."""
    
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
    
    # Get TOMORROW's date
    tomorrow = datetime.now() + timedelta(days=1)
    tomorrow_mmdd = f"{tomorrow.month:02d}-{tomorrow.day:02d}"
    
    logger.info(f"Checking for birthdays on {tomorrow.strftime('%B %d, %Y')}...")
    
    # Load members
    members = load_roster('data/members.xlsx')
    logger.info(f"Loaded {len(members)} members from roster")
    
    # Find members with birthdays tomorrow
    birthday_members = []
    for member in members:
        birthday = member.get('birthday')
        if birthday:
            try:
                if isinstance(birthday, datetime):
                    bday_mmdd = f"{birthday.month:02d}-{birthday.day:02d}"
                else:
                    bday_mmdd = str(birthday)
                
                if bday_mmdd == tomorrow_mmdd:
                    birthday_members.append(member)
            except:
                pass
    
    if not birthday_members:
        logger.warning(f"No members have birthdays on {tomorrow.strftime('%B %d, %Y')}")
        return False
    
    logger.info(f"Found {len(birthday_members)} member(s) with birthday(s) on {tomorrow.strftime('%B %d, %Y')}")
    
    try:
        # Load card config
        card_config = config.load_card_config('card_config.json')
        
        # Create output directory
        CARD_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        
        # Generate cards
        logger.info(f"Generating {len(birthday_members)} birthday card(s)...")
        card_files = []
        generated_count = 0
        
        for member in birthday_members:
            try:
                out_file = CARD_OUTPUT_DIR / f"card_{member['member_id']}.png"
                
                fonts_map = {}
                for font_ref, font_cfg in card_config.get('fonts', {}).items():
                    fonts_map[font_ref] = font_cfg.get('path', '')
                
                colors_map = card_config.get('colors', {})
                
                render_card(
                    member,
                    card_config,
                    str(out_file),
                    fonts_map,
                    colors_map,
                    base_path=Path('.')
                )
                
                card_files.append(out_file)
                generated_count += 1
                logger.info(f"✓ Generated card for {member['full_name']}")
                
            except Exception as e:
                logger.error(f"✗ Failed to generate card for {member['member_id']}: {e}")
        
        if generated_count == 0:
            logger.error("Failed to generate any cards")
            return False
        
        logger.info(f"Successfully generated {generated_count}/{len(birthday_members)} cards")
        
        # Handle both single string and list of recipients
        recipients = email['recipient']
        if isinstance(recipients, str):
            recipients = [recipients]
        
        # Log sender configuration
        sender_email = email.get('sender_email', 'default@trma.com.vn')
        use_alt = email.get('use_alt_sender', False)
        if use_alt:
            alt_email = email.get('alt_sender_email', sender_email)
            logger.info(f"📧 Sender: {email.get('sender_name', 'Team')} <{alt_email}> (alternative)")
        else:
            logger.info(f"📧 Sender: {email.get('sender_name', 'Team')} <{sender_email}>")
        
        subject = email['subject']
        
        # Get CC list
        cc_list = []
        if email.get('cc_birthday_people', False):
            logger.info("Extracting birthday people emails for CC...")
            cc_list = get_birthday_people_emails(card_files, members)
        
        # Add additional CC recipients
        additional_cc = email.get('additional_cc', [])
        if additional_cc:
            if isinstance(additional_cc, str):
                cc_list.append(additional_cc)
            else:
                cc_list.extend(additional_cc)
            logger.info(f"Adding additional CC recipients: {cc_list}")
        
        # Build HTML email body
        logger.info("Building email body with embedded cards...")
        
        cards_html = ""
        for card_file in card_files:
            with open(card_file, 'rb') as f:
                import base64
                card_base64 = base64.b64encode(f.read()).decode('utf-8')
                cards_html += f'<img src="data:image/png;base64,{card_base64}" style="max-width: 100%; height: auto; margin: 20px 0;"><br><br>'
        
        # Build HTML body
        html_body = f"""
<html>
<head>
<style>
    body {{ font-family: Segoe UI, Arial, sans-serif; {styling.get('body_color', '')} }}
    .container {{ max-width: 800px; margin: 0 auto; background: {styling.get('container_background', 'white')}; padding: 20px; border-radius: 10px; }}
    .header {{ text-align: center; color: {styling.get('header_color', '#000')}; font-size: {styling.get('header_size', '36px')}; margin-bottom: 20px; }}
    .content {{ text-align: center; color: {styling.get('body_color', '#000')}; margin: 20px 0; }}
    .card-container {{ text-align: center; margin: 20px 0; }}
    .footer {{ text-align: center; color: {styling.get('footer_color', '#666')}; font-size: 12px; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; }}
</style>
</head>
<body>
    <div class="container">
        <div class="header">{template.get('header_emoji', '🎂')} {template.get('header_text', 'Happy Birthday!')} {template.get('header_emoji', '🎂')}</div>
        <div class="content">
            <p>{template.get('greeting', 'Dear Birthday Person,')}</p>
            <p>{template.get('greeting_secondary', '')}</p>
        </div>
        <div class="card-container">
            {cards_html}
        </div>
        <div class="content">
            <p>{template.get('footer_text', '')}</p>
            <p>{template.get('closing', 'Warm regards,')}<br>{template.get('sender', '')}</p>
        </div>
        <div class="footer">{template.get('system_footer', 'Automated Birthday Greeting System')}</div>
    </div>
</body>
</html>
"""
        
        # Determine which sender email to use
        from_email = email.get('sender_email', 'default@trma.com.vn')
        if use_alt:
            from_email = email.get('alt_sender_email', from_email)
        
        # Send to all recipients
        logger.info(f"Sending to {len(recipients)} recipient(s)...")
        if cc_list:
            logger.info(f"With CC to: {', '.join(cc_list)}")
        
        all_sent = True
        
        for recipient_email in recipients:
            try:
                logger.info(f"Sending to: {recipient_email}...")
                
                success = send_email_via_outlook(
                    to_address=recipient_email,
                    subject=subject,
                    html_body=html_body,
                    attachments=None,
                    cc_addresses=cc_list if cc_list else None,
                    from_email=from_email
                )
                
                if success:
                    logger.info(f"✅ Email sent to {recipient_email}")
                else:
                    logger.error(f"❌ Failed to send to {recipient_email}")
                    all_sent = False
                    
            except Exception as e:
                logger.error(f"❌ Error sending to {recipient_email}: {e}")
                all_sent = False
        
        if all_sent:
            logger.info(f"✅ All {len(recipients)} email(s) sent successfully!")
            return True
        else:
            logger.error("❌ Some emails failed to send")
            return False
        
    except Exception as e:
        logger.error(f"❌ Workflow error: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return False

if __name__ == '__main__':
    success = main()
    exit(0 if success else 1)
