



📄 Automatic Invoice Generator (Excel Output)

This project generates a fully formatted invoice automatically.
You only enter the Client Name and Product Details, and the system handles everything else — calculations, taxes, totals, and Excel export.

Perfect for anyone who wants fast invoice creation without doing manual formatting.

🚀 Core Features

Auto-generate invoices in Excel (.xlsx)

Enter Client Name, Invoice Date, and Product List

Automatic calculation of:

Subtotal

CGST / SGST / IGST

Total Amount

Grand Total

Pre-designed Excel template with clean formatting

Supports multiple products

No manual Excel editing required

Ideal for freelancers, small businesses, and quick billing

🎥 Demo
<div align="center"> <video src="YOUR_VIDEO_URL_HERE" controls style="max-width: 100%; border-radius: 12px;"></video> </div>

<div align="center">
  <video src="https://github.com/user-attachments/assets/243af711-10a8-4796-80de-a8849f723e27](https://github.com/user-attachments/assets/74595535-7c38-406c-924e-7952ce76a8c5" 
         controls 
         style="max-width: 100%; border-radius: 12px;">
  </video>
</div>

📦 Tech Stack

Android / Kotlin

Apache POI (Excel generation)

Jetpack Compose / XML UI

Material Components

🧩 How It Works

User inputs:

Client Name

Invoice Date

Product Items (Name, Qty, Rate, HSN Code)

App calculates:

Line Amounts

Taxes

Totals

Excel sheet is created programmatically with:

Header section

Product table

Tax summary

Grand total

File is exported instantly to device storage.

📁 Project Structure
/invoice
  ├── ui/           # Input UI for invoice fields
  ├── model/        # Data models for products & invoice
  ├── excel/        # Apache POI Excel builder classes
  └── utils/        # Calculation helpers

🛠 Setup & Run

Clone the repo

Open in Android Studio

Run on device

Fill the form → click Generate Invoice

Excel file will be saved in storage

🎯 Purpose

This tool exists for one reason:
Make invoice creation stupidly simple and fast.

Upload details → Export → Done.

📄 License

MIT License – free for personal and commercial use.

If you want, I can also add:

✔ Screenshots section
✔ Download button
✔ Dark/light mode preview
✔ Badge icons (Android, Kotlin, License)
