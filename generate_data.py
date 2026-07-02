import pandas as pd
import random

random.seed(42)

merchant_data = {
    "Amazon": "Shopping",
    "Walmart": "Grocery",
    "Target": "Shopping",
    "Apple Store": "Electronics",
    "Netflix": "Entertainment",
    "Steam": "Gaming",
    "Crypto Exchange": "Crypto",
    "Airline": "Travel",
    "ATM Withdrawal": "Cash",
    "PayPal": "Shopping"
}

rows = []

for i in range(10000):
    merchant = random.choice(list(merchant_data.keys()))
    category = merchant_data[merchant]

    country = random.choices(
        ["USA", "Canada", "Mexico", "France", "Nigeria", "Russia"],
        weights=[70, 8, 7, 6, 5, 4],
        k=1
    )[0]

    payment_method = random.choices(
        ["Credit Card", "Debit Card", "Wire Transfer", "PayPal", "ATM"],
        weights=[45, 25, 5, 15, 10],
        k=1
    )[0]

    status = random.choices(
        ["Approved", "Declined"],
        weights=[92, 8],
        k=1
    )[0]

    amount = round(random.triangular(5, 10000, 120), 2)

    day = random.randint(1, 30)
    date = f"2026-06-{day:02d}"

    rows.append({
        "TransactionID": 1000 + i,
        "CustomerID": "C" + str(random.randint(1, 500)).zfill(3),
        "Date": date,
        "Merchant": merchant,
        "Category": category,
        "Amount": amount,
        "Country": country,
        "PaymentMethod": payment_method,
        "Status": status
    })

transactions = pd.DataFrame(rows)
transactions.to_csv("data/transactions.csv", index=False)


print("10,000 realistic transactions created!")