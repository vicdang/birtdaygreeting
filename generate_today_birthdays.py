"""
Generate birthday cards only for employees with birthdays today
"""
from birthday_bot.roster import load_roster
from birthday_bot import config
from birthday_bot.renderer import render_card
from pathlib import Path
from datetime import datetime

# Load members
members = load_roster('data/members.xlsx')

# Get today's date
today = datetime.now()
today_str = f"{today.month:02d}-{today.day:02d}"

# Find members with birthdays today
print(f"\nSearching for employees with birthdays on {today.strftime('%B %d, %Y')}...")
birthday_members = []

for member in members:
    birthday = member.get('birthday')
    if birthday:
        try:
            # Format birthday as MM-DD for comparison
            if isinstance(birthday, datetime):
                bday_str = f"{birthday.month:02d}-{birthday.day:02d}"
            else:
                # Parse string birthday
                bday_str = str(birthday)
            
            if bday_str == today_str:
                birthday_members.append(member)
        except:
            pass

if not birthday_members:
    print(f"No employees have birthdays today ({today.strftime('%B %d, %Y')})")
    print(f"No birthday cards to generate.")
    exit(0)

if birthday_members:
    print(f"\nGenerating {len(birthday_members)} birthday card(s)...")
    
    # Load config
    card_config = config.load_card_config('card_config.json')
    
    # Create output dir
    out_dir = Path('out/birthday_cards_today')
    out_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate cards
    success_count = 0
    error_count = 0
    
    for member in birthday_members:
        try:
            out_file = out_dir / f"card_{member['member_id']}.png"
            
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
            
            success_count += 1
            safe_name = member['full_name'].encode('ascii', 'replace').decode('ascii')[:40]
            print(f"[OK] {member['member_id']:10s} | {safe_name:40s} -> {out_file.name}")
        except Exception as e:
            error_count += 1
            print(f"[ER] {member['member_id']:10s} | {str(e)[:60]}")
    
    print(f"\n=== RESULTS ===")
    print(f"Success: {success_count}/{len(birthday_members)}")
    print(f"Errors: {error_count}/{len(birthday_members)}")
    print(f"Output directory: {out_dir}")
else:
    print("No birthday cards to generate today")
