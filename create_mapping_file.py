#!/usr/bin/env python3
"""
Helper script to create accounts_mapping.xlsx file

Run this script to create an example mapping file, then edit it with your actual data.
"""

import pandas as pd
from pathlib import Path

# Example data - REPLACE THIS WITH YOUR ACTUAL ACCOUNT DATA
example_data = {
    'account_number': [
        '1234567890',  # Replace with actual 10-digit account numbers
        '2345678901',
        '3456789012',
        '1234567890',  # Same person can have multiple accounts
    ],
    'guy_name': [
        'أحمد حسن',     # Replace with actual names (Arabic or English)
        'محمد علي',
        'سارة أحمد',
        'أحمد حسن',     # Same as first row - multiple accounts for same person
    ],
    'guy_email': [
        'ahmed.hassan@swd.bh',     # Replace with actual email addresses
        'mohammed.ali@swd.bh',
        'sara.ahmed@swd.bh',
        'ahmed.hassan@swd.bh',     # Same email for multiple accounts
    ]
}

# Create DataFrame
df = pd.DataFrame(example_data)

# Save to Excel
output_path = Path('config/accounts_mapping.xlsx')
output_path.parent.mkdir(parents=True, exist_ok=True)

df.to_excel(output_path, index=False, engine='openpyxl')

print(f"✅ Created mapping file: {output_path}")
print(f"\n📊 Example data ({len(df)} rows):")
print(df.to_string(index=False))
print(f"\n⚠️  IMPORTANT: Edit this file with your actual account data!")
print(f"   Open it in Excel and replace the example data with real account numbers.")
