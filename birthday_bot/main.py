"""Main CLI entrypoint for birthday bot."""

import logging
import sys
import argparse
from pathlib import Path
from datetime import datetime
from typing import Optional

from . import config, roster, renderer, mailer, state

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def setup_paths():
    """Set up base paths."""
    return Path(__file__).parent.parent


def run_daily_job(
    excel_path: str,
    config_path: str,
    date_override: Optional[str] = None,
    dry_run: bool = False
) -> int:
    """
    Run daily birthday job.
    
    Args:
        excel_path: Path to Excel roster
        config_path: Path to card config JSON
        date_override: Override date (YYYY-MM-DD)
        dry_run: Don't send emails
        
    Returns:
        Exit code (0=success, 1=failure)
    """
    try:
        base_path = setup_paths()
        
        # Load configuration
        logger.info("Loading configuration...")
        config.load_env()
        card_cfg = config.load_card_config(config_path)
        fonts_map = config.resolve_config_fonts(card_cfg, base_path)
        colors_map = config.resolve_config_colors(card_cfg)
        
        # Load roster
        logger.info(f"Loading roster from {excel_path}...")
        members = roster.load_roster(excel_path)
        members, invalid = roster.validate_member_photos(members, base_path)
        
        if invalid:
            logger.warning(f"Skipping {len(invalid)} members with invalid photos")
        
        # Determine target date
        if date_override:
            try:
                target_date = datetime.strptime(date_override, "%Y-%m-%d")
            except ValueError:
                logger.error(f"Invalid date format: {date_override}, use YYYY-MM-DD")
                return 1
        else:
            target_date = datetime.now()
        
        date_str = target_date.strftime("%Y-%m-%d")
        logger.info(f"Checking for birthdays on {date_str}")
        
        # Filter birthdays
        today_birthdays = roster.filter_birthdays_on_date(members, target_date)
        
        if not today_birthdays:
            logger.info("No birthdays today")
            return 0
        
        logger.info(f"Found {len(today_birthdays)} birthday(s)")
        
        # Load state tracker
        sent_log = state.SentLog()
        already_sent = sent_log.get_sent_today(date_str)
        
        # Render cards
        logger.info("Rendering cards...")
        card_paths = []
        members_to_send = []
        
        out_dir = base_path / "out" / date_str
        out_dir.mkdir(parents=True, exist_ok=True)
        
        for member in today_birthdays:
            if member['member_id'] in already_sent:
                logger.info(f"Skipping {member['member_id']} - already sent today")
                continue
            
            card_path = str(out_dir / f"card_{member['member_id']}.png")
            try:
                renderer.render_card(
                    member, card_cfg, card_path,
                    fonts_map, colors_map, base_path
                )
                card_paths.append(card_path)
                members_to_send.append(member)
            except renderer.RenderError as e:
                logger.error(f"Failed to render card for {member['member_id']}: {e}")
                return 1
        
        if not members_to_send:
            logger.info("All members already sent for today")
            return 0
        
        # Send emails
        if dry_run:
            logger.info("DRY RUN: Would send emails to:")
            for member in members_to_send:
                logger.info(f"  - {member['full_name']} <{member['email']}>")
            logger.info(f"Cards rendered to: {out_dir}")
            return 0
        
        logger.info("Sending emails...")
        email_mode = config.get_env('EMAIL_MODE', 'GROUP').upper()
        group_email = config.get_env('GROUP_EMAIL')
        
        try:
            mailer.send_birthday_email(
                members_to_send,
                card_paths,
                email_mode=email_mode,
                group_email=group_email
            )
        except mailer.MailerError as e:
            logger.error(f"Email sending failed: {e}")
            return 1
        
        # Mark as sent
        for member, card_path in zip(members_to_send, card_paths):
            sent_log.mark_sent(date_str, member['member_id'], member['email'])
        
        logger.info("Daily job completed successfully")
        return 0
    
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        return 1


def cmd_run(args):
    """Command: run"""
    return run_daily_job(
        args.excel,
        args.config,
        date_override=args.date,
        dry_run=args.dry_run
    )


def cmd_render(args):
    """Command: render"""
    return run_daily_job(
        args.excel,
        args.config,
        date_override=args.date,
        dry_run=True
    )


def cmd_test_email(args):
    """Command: test-email"""
    try:
        base_path = setup_paths()
        config.load_env()
        
        subject = "Test Email from Birthday Bot"
        html_body = """
        <html>
        <body>
            <h2>This is a test email</h2>
            <p>If you received this, SMTP configuration is working correctly!</p>
        </body>
        </html>
        """
        
        attachments = []
        if args.attach:
            attachments = [{'path': args.attach, 'filename': Path(args.attach).name}]
        
        mailer.send_email(args.to, subject, html_body, attachments=attachments)
        logger.info(f"Test email sent to {args.to}")
        return 0
    
    except Exception as e:
        logger.error(f"Test email failed: {e}", exc_info=True)
        return 1


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Birthday Card Generator Bot',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m birthday_bot.main run --excel members.xlsx
  python -m birthday_bot.main render --excel members.xlsx --date 2026-02-25 --dry-run
  python -m birthday_bot.main test-email --to admin@example.com
        """
    )
    
    parser.add_argument(
        '--config',
        default='card_config.json',
        help='Path to card config (default: card_config.json)'
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Command to run')
    
    # run command
    run_parser = subparsers.add_parser('run', help='Run daily birthday job')
    run_parser.add_argument('--excel', required=True, help='Path to Excel roster')
    run_parser.add_argument('--date', help='Override date (YYYY-MM-DD)')
    run_parser.add_argument('--dry-run', action='store_true', help='Render only, no email')
    run_parser.set_defaults(func=cmd_run)
    
    # render command
    render_parser = subparsers.add_parser('render', help='Render cards (dry-run)')
    render_parser.add_argument('--excel', required=True, help='Path to Excel roster')
    render_parser.add_argument('--date', help='Override date (YYYY-MM-DD)')
    render_parser.set_defaults(func=cmd_render)
    
    # test-email command
    test_parser = subparsers.add_parser('test-email', help='Send test email')
    test_parser.add_argument('--to', required=True, help='Recipient email')
    test_parser.add_argument('--attach', help='File to attach')
    test_parser.set_defaults(func=cmd_test_email)
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return 1
    
    return args.func(args)


if __name__ == '__main__':
    sys.exit(main())
