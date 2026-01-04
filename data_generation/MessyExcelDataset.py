import pandas as pd
import numpy as np
import random
from faker import Faker

fake = Faker()
np.random.seed(42)
random.seed(42)

rows = 300

# Generate base data
data = {
    "customer_id": [random.randint(1000, 1100) for _ in range(rows)],  # duplicates
    "name": [fake.name() if random.random() > 0.05 else None for _ in range(rows)],  # missing
    "age": [
        random.choice([random.randint(18, 80), None, -5, 150])  # invalid ages
        for _ in range(rows)
    ],
    "email": [
        fake.email() if random.random() > 0.1 else fake.word()  # broken emails
        for _ in range(rows)
    ],
    "signup_date": [
        random.choice([
            fake.date_this_decade().strftime("%Y-%m-%d"),
            fake.date_this_decade().strftime("%d/%m/%Y"),
            fake.date_this_decade().strftime("%m-%d-%Y"),
            None
        ])
        for _ in range(rows)
    ],
    "country": [
        random.choice([
            "USA", "United States", "us", "U.S.A",
            "UK", "United Kingdom", "uk",
            "India", "IND", "in",
            None
        ])
        for _ in range(rows)
    ],
    "purchase_amount": [
        round(random.choice([random.uniform(5, 500), -50, 99999]), 2)  # outliers & negatives
        for _ in range(rows)
    ],
    "membership_level": [
        random.choice(["Gold", "gold", "GOLD", "Silver", "silver", "Bronze", "", None])
        for _ in range(rows)
    ],
}

df = pd.DataFrame(data)

# Introduce full duplicate rows
df = pd.concat([df, df.sample(20, random_state=1)], ignore_index=True)

# Shuffle rows
df = df.sample(frac=1, random_state=99).reset_index(drop=True)

# Save to Excel
df.to_excel("messy_dataset.xlsx", index=False)

print("messy_dataset.xlsx created successfully!")
