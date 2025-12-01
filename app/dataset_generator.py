import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import os

# Create directory for monthly Excel files
os.makedirs('monthly_billing_data', exist_ok=True)

# Sample data configuration
customer_names = [
    'Rajesh Patel', 'Priya Shah', 'Amit Kumar', 'Neha Desai', 'Vikram Singh',
    'Anjali Mehta', 'Karan Sharma', 'Ritu Joshi', 'Manish Gupta', 'Pooja Trivedi',
    'Sanjay Rao', 'Divya Nair', 'Rahul Verma', 'Sneha Iyer', 'Arjun Reddy',
    'Kavita Bhatia', 'Nikhil Agarwal', 'Meera Kulkarni', 'Rohan Kapoor', 'Shruti Pandey'
]

jewelry_items = [
    'Gold Necklace', 'Gold Ring', 'Gold Earrings', 'Gold Bracelet', 'Gold Chain',
    'Silver Necklace', 'Silver Ring', 'Silver Earrings', 'Silver Anklet',
    'Diamond Ring', 'Diamond Earrings', 'Diamond Pendant', 'Diamond Bracelet',
    'Gold Bangle', 'Silver Bangle', 'Nose Pin', 'Mangalsutra', 'Gold Pendant'
]

months = ['January', 'February', 'March', 'April', 'May', 'June', 
          'July', 'August', 'September', 'October', 'November', 'December']

def generate_monthly_data(month_num, year=2024):
    """Generate billing data for a specific month"""
    month_name = months[month_num - 1]
    
    # Random number of transactions per month (50-80)
    num_transactions = random.randint(50, 80)
    
    data = {
        'Bill_No': [],
        'Date': [],
        'Customer_Name': [],
        'Contact_Number': [],
        'Item_Name': [],
        'Quantity': [],
        'Weight_Grams': [],
        'Rate_Per_Gram': [],
        'Making_Charges': [],
        'Total_Amount': [],
        'Payment_Mode': []
    }
    
    for i in range(num_transactions):
        # Generate date within the month
        day = random.randint(1, 28 if month_num == 2 else 30 if month_num in [4,6,9,11] else 31)
        date = datetime(year, month_num, day).strftime('%Y-%m-%d')
        
        # Generate bill number
        bill_no = f"JB{year}{month_num:02d}{i+1:04d}"
        
        # Random customer
        customer = random.choice(customer_names)
        contact = f"+91 {''.join([str(random.randint(0,9)) for _ in range(10)])}"
        
        # Random item
        item = random.choice(jewelry_items)
        quantity = random.randint(1, 3)
        
        # Weight and pricing based on item type
        if 'Gold' in item:
            weight = round(random.uniform(5, 50), 2)
            rate = round(random.uniform(5500, 6000), 2)
            making_charges = round(weight * random.uniform(400, 600), 2)
        elif 'Silver' in item:
            weight = round(random.uniform(10, 100), 2)
            rate = round(random.uniform(70, 85), 2)
            making_charges = round(weight * random.uniform(50, 100), 2)
        elif 'Diamond' in item:
            weight = round(random.uniform(2, 20), 2)
            rate = round(random.uniform(3000, 5000), 2)
            making_charges = round(weight * random.uniform(1000, 2000), 2)
        else:
            weight = round(random.uniform(5, 30), 2)
            rate = round(random.uniform(5000, 6000), 2)
            making_charges = round(weight * random.uniform(300, 500), 2)
        
        # Calculate total
        total = round((weight * rate * quantity) + making_charges, 2)
        
        # Payment mode
        payment = random.choice(['Cash', 'Card', 'UPI', 'Bank Transfer'])
        
        # Append to data
        data['Bill_No'].append(bill_no)
        data['Date'].append(date)
        data['Customer_Name'].append(customer)
        data['Contact_Number'].append(contact)
        data['Item_Name'].append(item)
        data['Quantity'].append(quantity)
        data['Weight_Grams'].append(weight)
        data['Rate_Per_Gram'].append(rate)
        data['Making_Charges'].append(making_charges)
        data['Total_Amount'].append(total)
        data['Payment_Mode'].append(payment)
    
    return pd.DataFrame(data)

# Generate data for all 12 months
print("Generating monthly billing data for 2024...")
for month_num in range(1, 13):
    df = generate_monthly_data(month_num)
    
    # Add some realistic data quality issues for preprocessing demonstration
    # Random missing values (2-3%)
    if random.random() > 0.5:
        missing_idx = random.sample(range(len(df)), k=random.randint(1, 3))
        df.loc[missing_idx, 'Contact_Number'] = np.nan
    
    # Random duplicates (1-2 rows)
    if random.random() > 0.7:
        dup_idx = random.sample(range(len(df)), k=random.randint(1, 2))
        df = pd.concat([df, df.iloc[dup_idx]], ignore_index=True)
    
    # Save to Excel
    filename = f'monthly_billing_data/{months[month_num-1]}_2024.xlsx'
    df.to_excel(filename, index=False)
    print(f"✓ Created {filename} with {len(df)} records")

print("\n✅ All monthly Excel files generated successfully!")
print(f"📁 Files saved in: monthly_billing_data/")
print("\nSample data from January 2024:")
jan_df = pd.read_excel('monthly_billing_data/January_2024.xlsx')
print(jan_df.head())