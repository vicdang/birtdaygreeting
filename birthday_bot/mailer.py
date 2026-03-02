"""Email sending with SMTP and Outlook."""

import logging
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from typing import List, Dict, Any, Optional
from datetime import datetime
from PIL import Image

from .config import get_env, get_required_env, ConfigError

logger = logging.getLogger(__name__)


class MailerError(Exception):
    """Email sending error."""
    pass


def build_collage(
    card_paths: List[str],
    output_path: str,
    columns: int = 2,
    padding: int = 20
) -> str:
    """
    Build a grid collage from card images.
    
    Args:
        card_paths: List of card image paths
        output_path: Output collage path
        columns: Number of columns in grid
        padding: Padding between cards in pixels
        
    Returns:
        Path to collage image
    """
    if not card_paths:
        raise MailerError("No cards to collage")
    
    try:
        # Load all images
        images = []
        for path in card_paths:
            img = Image.open(path)
            if img.mode != 'RGBA':
                img = img.convert('RGBA')
            images.append(img)
        
        # Calculate grid layout
        card_width, card_height = images[0].size
        rows = (len(images) + columns - 1) // columns
        
        # Create collage canvas
        collage_width = columns * card_width + (columns - 1) * padding
        collage_height = rows * card_height + (rows - 1) * padding
        
        collage = Image.new('RGBA', (collage_width, collage_height), (0, 0, 0, 0))
        
        # Paste images
        for i, img in enumerate(images):
            row = i // columns
            col = i % columns
            x = col * (card_width + padding)
            y = row * (card_height + padding)
            collage.paste(img, (x, y), img)
        
        # Save
        collage.save(output_path, 'PNG')
        logger.info(f"Created collage: {output_path}")
        
        return output_path
    
    except Exception as e:
        raise MailerError(f"Failed to create collage: {e}")


def send_email(
    to_address: str,
    subject: str,
    html_body: str,
    text_body: Optional[str] = None,
    attachments: Optional[List[Dict[str, str]]] = None
) -> bool:
    """
    Send email via SMTP.
    
    Args:
        to_address: Recipient email address
        subject: Email subject
        html_body: HTML email body
        text_body: Plain text fallback (optional)
        attachments: List of {path, filename, mimetype} dicts
        
    Returns:
        True if sent successfully
        
    Raises:
        MailerError: If sending fails
    """
    try:
        # Get SMTP config
        smtp_host = get_required_env('SMTP_HOST')
        smtp_port = int(get_env('SMTP_PORT', '587'))
        smtp_user = get_required_env('SMTP_USER')
        smtp_pass = get_required_env('SMTP_PASS')
        sender_name = get_env('SENDER_NAME', 'Birthday Bot')
        sender_email = get_env('SENDER_EMAIL', smtp_user)
        
        # Check if alternative sender should be used
        alt_sender_email = get_env('ALT_SENDER_EMAIL')
        use_alt = get_env('USE_ALT_SENDER', 'false').lower() == 'true'
        if use_alt and alt_sender_email:
            sender_email = alt_sender_email
            logger.info(f"Using alternative sender: {sender_email}")
    except ConfigError as e:
        raise MailerError(f"Missing SMTP configuration: {e}")
    
    try:
        # Create message
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = f"{sender_name} <{sender_email}>"
        msg['To'] = to_address
        
        # Add text part (fallback)
        if text_body:
            msg.attach(MIMEText(text_body, 'plain'))
        else:
            # Strip HTML tags for text version
            import re
            text = re.sub('<[^<]+?>', '', html_body)
            msg.attach(MIMEText(text, 'plain'))
        
        # Add HTML part
        msg.attach(MIMEText(html_body, 'html'))
        
        # Add attachments
        if attachments:
            for att in attachments:
                _attach_file(msg, att['path'], att.get('filename', Path(att['path']).name))
        
        # Send
        if smtp_port == 465:
            # Port 465 uses implicit TLS (SMTP_SSL)
            with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        else:
            # Port 587 uses explicit TLS (SMTP + starttls)
            with smtplib.SMTP(smtp_host, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_pass)
                server.send_message(msg)
        
        logger.info(f"Email sent to {to_address}")
        return True
    
    except smtplib.SMTPException as e:
        raise MailerError(f"SMTP error: {e}")
    except Exception as e:
        raise MailerError(f"Failed to send email: {e}")


def _attach_file(msg: MIMEMultipart, file_path: str, filename: str):
    """
    Attach file to email message.
    
    Args:
        msg: MIMEMultipart message
        file_path: Path to file to attach
        filename: Display filename in email
    """
    file_path = Path(file_path)
    
    if not file_path.exists():
        logger.warning(f"Attachment not found: {file_path}")
        return
    
    try:
        # Detect MIME type
        if file_path.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif']:
            with open(file_path, 'rb') as f:
                part = MIMEImage(f.read())
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(part)
        else:
            # Default to application/octet-stream
            with open(file_path, 'rb') as f:
                data = f.read()
            from email.mime.base import MIMEBase
            from email import encoders
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(part)
    
    except Exception as e:
        logger.error(f"Failed to attach {file_path}: {e}")


def send_birthday_email(
    members: List[Dict[str, Any]],
    card_paths: List[str],
    email_mode: str = "GROUP",
    group_email: Optional[str] = None
) -> bool:
    """
    Send birthday email(s).
    
    Args:
        members: List of birthday members
        card_paths: List of card image paths (same order as members)
        email_mode: "GROUP" or "PERSON"
        group_email: Recipient if GROUP mode (optional, defaults to env)
        
    Returns:
        True if all emails sent successfully
        
    Raises:
        MailerError: If any email fails
    """
    if not members or not card_paths:
        raise MailerError("No members or cards to send")
    
    if len(members) != len(card_paths):
        raise MailerError("Member and card path count mismatch")
    
    try:
        subject = get_env('EMAIL_SUBJECT', 'Happy Birthday 🎉')
        
        if email_mode == "GROUP":
            return _send_group_email(members, card_paths, subject, group_email)
        else:
            return _send_person_emails(members, card_paths, subject)
    
    except MailerError:
        raise
    except Exception as e:
        raise MailerError(f"Email sending failed: {e}")


def _send_group_email(
    members: List[Dict[str, Any]],
    card_paths: List[str],
    subject: str,
    group_email: Optional[str] = None
) -> bool:
    """Send one email to group address with all cards."""
    if not group_email:
        group_email = get_required_env('GROUP_EMAIL')
    
    # Build member list for email
    member_names = [m['full_name'] for m in members]
    member_list = '<br>'.join(f"🎂 {name}" for name in member_names)
    
    # Build HTML body
    html_body = f"""
    <html>
    <body>
        <h2>Happy Birthday! 🎉</h2>
        <p>Today we celebrate the birthday of:</p>
        <div style="margin: 20px 0;">
            {member_list}
        </div>
        <p>See attached birthday cards below!</p>
    </body>
    </html>
    """
    
    # Prepare attachments
    attachments = [
        {'path': path, 'filename': f'card_{i}.png'}
        for i, path in enumerate(card_paths)
    ]
    
    # Add collage if enabled
    if get_env('ATTACH_COLLAGE', '').lower() == 'true':
        try:
            collage_path = 'out/collage.png'
            build_collage(card_paths, collage_path)
            attachments.insert(0, {'path': collage_path, 'filename': 'birthday_collage.png'})
        except Exception as e:
            logger.warning(f"Failed to create collage: {e}")
    
    return send_email(group_email, subject, html_body, attachments=attachments)


def _send_person_emails(
    members: List[Dict[str, Any]],
    card_paths: List[str],
    subject: str
) -> bool:
    """Send individual emails to each birthday person."""
    all_sent = True
    
    for member, card_path in zip(members, card_paths):
        try:
            to_address = member['email']
            name = member['full_name']
            
            html_body = f"""
            <html>
            <body>
                <h2>Happy Birthday, {name}! 🎉</h2>
                <p>Wishing you an amazing day filled with joy and celebration!</p>
                <p>Your birthday card is attached.</p>
            </body>
            </html>
            """
            
            attachments = [{'path': card_path, 'filename': f'{member["member_id"]}_card.png'}]
            
            send_email(to_address, subject, html_body, attachments=attachments)
        
        except MailerError as e:
            logger.error(f"Failed to send email to {member['email']}: {e}")
            all_sent = False
    
    return all_sent


def send_email_via_outlook(
    to_address: str,
    subject: str,
    html_body: str = "",
    text_body: str = "",
    attachments: List[Dict[str, str]] = None,
    cc_addresses: List[str] = None,
    from_email: str = None
) -> bool:
    """
    Send email using Outlook COM interface.
    
    Requires: Outlook to be installed and configured on Windows
    Advantage: Uses Outlook's already-configured credentials
    
    Args:
        to_address: Recipient email address
        subject: Email subject
        html_body: HTML email body
        text_body: Plain text fallback (optional)
        attachments: List of {"path": file_path, "filename": display_name}
        cc_addresses: List of email addresses to CC
        from_email: Sender email address (optional, uses default if not specified)
        
    Returns:
        True if email sent successfully
        
    Raises:
        MailerError: If Outlook is not available or sending fails
    """
    try:
        import win32com.client
    except ImportError:
        raise MailerError("pywin32 not installed. Install with: pip install pywin32")
    
    try:
        # Get Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Create mail item
        mail = outlook.CreateItem(0)  # 0 = MailItem
        
        # Set basic properties
        mail.To = to_address
        mail.Subject = subject
        mail.HTMLBody = html_body if html_body else text_body
        
        # Set sender if specified
        if from_email:
            try:
                # Try to find the account matching the email address
                accounts = outlook.Session.Accounts
                account_found = False
                for account in accounts:
                    if account.SmtpAddress.lower() == from_email.lower():
                        mail.SendUsingAccount = account
                        logger.info(f"Set sender account: {from_email}")
                        account_found = True
                        break
                
                if not account_found:
                    # If exact match not found, log warning but continue
                    logger.warning(f"Outlook account not found for {from_email}, using default account")
            except Exception as e:
                logger.warning(f"Could not set sender account: {e}, using default account")
        
        # Add CC addresses if provided
        if cc_addresses:
            mail.CC = '; '.join(cc_addresses)
            logger.info(f"CC'd to: {', '.join(cc_addresses)}")
        
        # Attach files if provided
        if attachments:
            for att in attachments:
                att_path = Path(att['path'])
                if att_path.exists():
                    mail.Attachments.Add(str(att_path.absolute()))
                    logger.info(f"Attached: {att['filename']}")
                else:
                    logger.warning(f"Attachment not found: {att_path}")
        
        # Send email
        mail.Send()
        
        logger.info(f"Email sent via Outlook to {to_address}")
        return True
    
    except Exception as e:
        raise MailerError(f"Failed to send email via Outlook: {e}")
