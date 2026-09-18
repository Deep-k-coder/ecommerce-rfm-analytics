"""
generate_dataset.py
Generates a realistic, highly representative e-commerce transactional dataset
conforming to the canonical Online Retail schema, complete with real-world edge cases.
"""
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

def generate_synthetic_retail_data(n_rows: int = 15000, random_seed: int = 42) -> pd.DataFrame:
    np.random.seed(random_seed)
    
    # Base catalog items (StockCode, Description, UnitPrice)
    catalog = [
        ("85123A", "WHITE HANGING HEART T-LIGHT HOLDER", 2.55),
        ("71053", "WHITE METAL LANTERN", 3.39),
        ("84406B", "CREAM CUPID HEARTS COAT HANGER", 2.75),
        ("84029G", "KNITTED UNION FLAG HOT WATER BOTTLE", 3.39),
        ("84029E", "RED WOOLLY HOTTIE WHITE HEART.", 3.39),
        ("22752", "SET 7 BABUSHKA NESTING BOXES", 7.65),
        ("21730", "GLASS STAR FROSTED T-LIGHT HOLDER", 4.25),
        ("22633", "HAND WARMER UNION JACK", 1.85),
        ("22632", "HAND WARMER RED POLKA DOT", 1.85),
        ("84879", "ASSORTED COLOUR BIRD ORNAMENT", 1.69),
        ("22745", "POPPY'S PLAYHOUSE BEDROOM", 2.10),
        ("22748", "POPPY'S PLAYHOUSE KITCHEN", 2.10),
        ("22749", "POPPY'S PLAYHOUSE LIVINGROOM", 2.10),
        ("22310", "IVORY KNITTED MUG COSY", 1.65),
        ("84969", "BOX OF 6 ASSORTED COLOUR TEASPOONS", 4.25),
        ("22086", "PAPER CHAIN KIT 50'S CHRISTMAS", 2.55),
        ("22630", "DOORMAT KEEP CALM AND COME IN", 6.75),
        ("47566", "PARTY BUNTING", 4.65),
        ("20725", "LUNCH BAG RED RETROSPOT", 1.65),
        ("20727", "LUNCH BAG BLACK SKULL.", 1.65),
        ("20728", "LUNCH BAG CARS BLUE", 1.65),
        ("22382", "LUNCH BAG SPACEBOY DESIGN", 1.65),
        ("22383", "LUNCH BAG SUKI DESIGN", 1.65),
        ("22384", "LUNCH BAG PINK POLKADOT", 1.65),
        ("20726", "LUNCH BAG WOODLAND", 1.65),
        ("22386", "JUMBO BAG PINK POLKADOT", 1.95),
        ("85099B", "JUMBO BAG RED RETROSPOT", 1.95),
        ("21931", "JUMBO STORAGE BAG SUKI", 1.95),
        ("85099F", "JUMBO BAG STRAWBERRY", 1.95),
        ("22411", "JUMBO SHOPPER VINTAGE RED PAISLEY", 1.95)
    ]
    
    # 800 distinct registered customer IDs + None for guests
    customer_ids = [float(cid) for cid in range(12346, 13146)]
    
    # Observation window: Jan 01 2023 to Dec 10 2023
    start_date = datetime(2023, 1, 1, 8, 0, 0)
    end_date = datetime(2023, 12, 10, 18, 0, 0)
    time_delta = end_date - start_date

    # Behavior profiles for customers:
    # 15% high-frequency champions, 25% loyal regular, 30% occasional, 30% one-off/dormant
    cust_weights = np.random.pareto(a=1.5, size=len(customer_ids))
    cust_weights /= cust_weights.sum()

    records = []
    invoice_seq = 536000
    
    # Generate batch of orders
    current_time = start_date
    while len(records) < n_rows:
        invoice_seq += 1
        invoice_no = str(invoice_seq)
        
        # 12% probability of missing CustomerID (guest checkouts)
        if np.random.rand() < 0.12:
            cid = np.nan
        else:
            cid = np.random.choice(customer_ids, p=cust_weights)
            
        # Timestamp distribution with holiday clustering towards Q4
        sec_offset = int(np.random.beta(a=2.0, b=1.2) * time_delta.total_seconds())
        inv_date = start_date + timedelta(seconds=sec_offset)
        
        # 4% chance of being a canceled invoice
        is_cancellation = np.random.rand() < 0.04
        if is_cancellation:
            invoice_no = f"C{invoice_no}"
            
        # Number of line items in this invoice (1 to 6)
        n_items = np.random.randint(1, 7)
        chosen_items = np.random.choice(len(catalog), size=n_items, replace=False)
        
        for idx in chosen_items:
            stock_code, desc, base_price = catalog[idx]
            
            # Quantity logic
            if is_cancellation:
                qty = -int(np.random.randint(1, 6))
                price = base_price
            else:
                # 0.5% extreme wholesale bulk outlier (e.g., 500 to 2000 units) to test IQR
                if np.random.rand() < 0.005:
                    qty = int(np.random.randint(400, 1500))
                else:
                    qty = int(np.random.choice([1, 2, 3, 4, 6, 12, 24], p=[0.35, 0.25, 0.15, 0.1, 0.08, 0.05, 0.02]))
                
                # 0.2% price anomaly (zero or extreme price)
                if np.random.rand() < 0.002:
                    price = 0.0 if np.random.rand() < 0.5 else 450.00
                else:
                    price = round(base_price * np.random.uniform(0.95, 1.05), 2)
            
            records.append({
                "InvoiceNo": invoice_no,
                "StockCode": stock_code,
                "Description": desc,
                "Quantity": qty,
                "InvoiceDate": inv_date.strftime("%Y-%m-%d %H:%M:%S"),
                "UnitPrice": price,
                "CustomerID": cid
            })
            
            if len(records) >= n_rows:
                break

    df = pd.DataFrame(records)
    return df

if __name__ == "__main__":
    df = generate_synthetic_retail_data(n_rows=16000)
    output_path = "/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics/data/raw/online_retail.csv"
    df.to_csv(output_path, index=False)
    print(f"[DATA GENERATOR] Successfully generated {len(df):,} transactions at {output_path}")
    print(f"Columns: {list(df.columns)}")
    print(f"Missing CustomerIDs: {df['CustomerID'].isna().sum():,}")
    print(f"Canceled Orders ('C'): {df['InvoiceNo'].str.startswith('C').sum():,}")
    print(f"Negative Quantities: {(df['Quantity'] < 0).sum():,}")
