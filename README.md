# Aurella OMS — Setup Guide

## 1. Google Sheet Setup

Create a new Google Sheet with these 3 sheets (tabs):
- `Orders`
- `OrderItems`  
- `ActivityLog`

The Apps Script will auto-create headers on first use.

**Copy your Sheet ID** from the URL:
`https://docs.google.com/spreadsheets/d/` **← THIS PART →** `/edit`

## 2. Apps Script Setup

1. In your Google Sheet → **Extensions → Apps Script**
2. Delete the default `Code.gs` content
3. Paste the contents of `Code.gs` from this folder
4. Replace `const SHEET_ID = '';` with your actual Sheet ID
5. Save (Ctrl+S)

## 3. Deploy the Web App

1. Click **Deploy → New Deployment**
2. Type: **Web App**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. Click **Deploy**
6. Copy the deployment URL (starts with `https://script.google.com/macros/s/...`)

> ⚠️ Every time you edit `Code.gs`, create a **New Deployment** (not update) to get a fresh URL.

## 4. Connect the OMS

1. Open `index.html` in your browser (or deploy to GitHub Pages)
2. Click **⚙ Configure API** in the sidebar
3. Paste your Apps Script URL
4. Click **Save & Test**

You should see the status indicator turn green.

## 5. GitHub Pages Deployment

1. Create a GitHub repo (e.g. `aurella-oms`)
2. Upload `index.html`
3. Go to Settings → Pages → Source: `main` branch, root `/`
4. Your OMS will be live at `https://yourusername.github.io/aurella-oms/`

---

## Column Reference

### Orders Sheet
| Column | Description |
|--------|-------------|
| OrderID | Primary key (AUR-0001) |
| CustomerName | Full name |
| Phone | Phone number |
| Address | Street address |
| Landmark | Landmark |
| Pincode | PIN code |
| DeliveryType | Handover / Pickup / Shipping |
| Courier | Courier name (if Shipping) |
| AWB | Airway Bill number |
| PaymentStatus | Paid / Not Paid |
| OrderStatus | Confirmed / In Progress / Ready / In Transit / Delivered / RTO - In Transit / Cancelled |
| CreatedAt | ISO timestamp |
| UpdatedAt | ISO timestamp |

### OrderItems Sheet
| Column | Description |
|--------|-------------|
| OrderID | Foreign key |
| ProductName | Free text product name |
| Quantity | Number |

### ActivityLog Sheet
| Column | Description |
|--------|-------------|
| OrderID | Foreign key |
| Description | What changed |
| Timestamp | ISO timestamp |

---

## Bulk CSV Formats

### Bulk Create
```
Order ID,Customer Name,Phone Number,Address,Landmark,Pincode,Delivery Type,Payment Status,Order Status,Product Name,Quantity
AUR-0010,Customer Name,9876543210,Address,Landmark,560001,Shipping,Not Paid,Confirmed,Chunky Bag,1
AUR-0010,Customer Name,9876543210,Address,Landmark,560001,Shipping,Not Paid,Confirmed,Hair Clip,2
```

### Bulk Edit (leave blank to keep existing value)
```
Order ID,Customer Name,Phone Number,Address,Landmark,Pincode,Delivery Type,Courier,AWB Number,Payment Status,Order Status
AUR-0001,,,,,,,,,Paid,
AUR-0002,,,,,,,Delhivery,DEL123456,,In Transit
```
