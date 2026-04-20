# **🔢 Excel Number to Words Converter (SpellNumber Add-in)**

Welcome to the **Excel Number to Words Converter** repository! 
This project provides a powerful **Excel Add-in (.xlam)** that enables users to convert numeric values into **readable words** directly inside Excel using a simple formula.

---

## **📖 Project Overview**

Excel does not provide a built-in function to convert numbers into words.
This project solves that limitation by introducing a **custom VBA-based global function**.

### ✨ Key Capabilities:

✅ Works across **all Excel files** (via Add-in)  
✅ Converts **integers, decimals, and negative numbers**  
✅ Supports **Indian Number System (Thousand, Lakh, Crore)**
✅ Handles **comma-separated values automatically**  
✅ Clean and reusable implementation

---

## **⚙️ Function Usage**

Once installed, use:

```excel
=SpellNumber(cell_reference)
```

---

## **🧪 Example Outputs**

| Input    | Output                                   |
| -------- | ---------------------------------------- |
| 123      | One Hundred Twenty Three                 |
| 123.45   | One Hundred Twenty Three Point Four Five |
| -500     | Minus Five Hundred                       |
| 1,25,000 | One Lakh Twenty Five Thousand            |
| 0        | Zero                                     |

---

# **🧩 Installation Guide (For All Users)**

> This section is for beginners. Just follow the steps—no coding needed.

---

## **1️⃣ Download the Add-in File**

* Download:

```
SpellNumber.xlam
```

(from this repository)

---

## **2️⃣ Open Microsoft Excel**

* Open Excel normally
  📸 Example:\

---

## **3️⃣ Go to Add-ins Menu**

* Click:

```
File → Options → Add-ins
```

---

## **4️⃣ Open Excel Add-ins**

At the bottom:

```
Manage: Excel Add-ins → Go
```

---

## **5️⃣ Install the Add-in**

* Click **Browse**
* Select the downloaded file:

```
SpellNumber.xlam
```

* Click **OK**
* Tick the checkbox ✔
* Click **OK**

📸 Example:\

---

## **6️⃣ Enable Macros (Important)**

If prompted:

👉 Click **Enable Content**

---

## **7️⃣ Start Using**

Now in any Excel file:

```excel
=SpellNumber(A1)
```

📸 Example:\

---

# **🛠️ Implementation Flow (For Developers / Advanced Users)**

> 💡 This section is for users who want to **build or customize the add-in**

---

## 🔹 Step 1: Open VBA Editor

* Press:

```
Alt + F11
```

📸\

---

## 🔹 Step 2: Insert Module & Paste Code

* Go to:

```
Insert → Module
```

* Paste the SpellNumber VBA code

📸\

---

## 🔹 Step 3: Save as Macro Workbook (.xlsm)

* File → Save As
* Choose:

```
Excel Macro-Enabled Workbook (*.xlsm)
```

📸\

---

## 🔹 Step 4: Convert to Add-in (.xlam)

* File → Save As
* Choose:

```
Excel Add-in (*.xlam)
```

📸\

---

## 🔹 Step 5: Install Add-in

* Same steps as Installation Guide above

---

# **⚠️ Important Notes**

* Macros must be enabled
* File must be `.xlam` (not `.xlsm`)
* Restart Excel if function doesn’t appear
* Ensure add-in is active under **Excel Add-ins**

---

```
Excel-SpellNumber-Add-In/
│
├── Screenshots/
│   ├── step1_vba.png
│   ├── step2_code.png
│   ├── step3_xlsm.png
│   ├── step4_xlam.png
│   ├── step5_addin.png
│   └── step6_usage.png
│
├── SpellNumber.xlam
├── SpellNumber.xlsm
├── VBA-Code.txt
├── README.md
└── LICENSE
```

---

# **📦 Sharing with Others**

* Share the `.xlam` file
* Each user installs it once
* Recommended: store in a **shared network folder**

---

# **🛠️ Tools & Technologies Used**

* Microsoft Excel (VBA)
* Visual Basic for Applications (VBA)

---

# **🔮 Future Enhancements**

🚀 Currency format (₹ Rupees / Paise)
🚀 Excel Ribbon Button (no formula needed)
🚀 Multi-language support
🚀 PHP integration for backend systems

---

# **🔗 Connect with Me**

Hi! I'm **Shivam Singh**, a **Data Analyst** passionate about building automation tools and data-driven solutions.

🚀 **Stay tuned for more automation & data projects!**
