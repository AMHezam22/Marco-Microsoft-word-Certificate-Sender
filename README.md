# 🎓 Certificate Generator Macro (Word + Excel + Outlook)

This repository contains a VBA macro script (`GenerateCertificates.bas`) that automates the process of generating and emailing personalized PDF certificates.

## ✨ Features

- Reads recipient names and emails from an Excel sheet
- Replaces a placeholder (`[NAME]`) inside a **Word text box**
- Keeps text formatting (e.g., bold, font size)
- Exports personalized certificates as PDFs
- Sends each certificate by email via Microsoft Outlook
- Reverts the placeholder after each generation

## 📂 Requirements

- Microsoft Word (with a `.docm` document template)
- Microsoft Excel (with a `.xlsx` file containing name/email)
- Microsoft Outlook (for sending emails)
- VBA enabled in your Office installation

## 🛠 How to Use

1. **Download this repository** or copy `GenerateCertificates.bas`
2. Open your `.docm` Word certificate template
3. Press `Alt + F11` to open the VBA editor
4. Go to `File > Import File...`, and select `GenerateCertificates.bas`
5. Save the Word file as **`.docm`**
6. Prepare an Excel file:
    - Column A: Full Name  
    - Column B: Email Address  
7. Run the macro `GenerateAndEmailCertificates` from Word

## 📝 Excel Format Example

| Name            | Email              |
|-----------------|--------------------|
| John Smith      | john@example.com   |
| Sarah Johnson   | sarah@example.com  |

## 📌 Placeholder in Word

Make sure the Word document includes the text `[NAME]` **inside a text box**, and format it however you like (e.g., bold). The macro will replace that text and retain the formatting.

## 🔐 Macro Security

If macros are disabled on your system:
- Open Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
- Choose **"Disable all macros with notification"**
- Click **"Enable Content"** when prompted

## 📦 File Structure
📁 your-repo/
├─ GenerateCertificates.bas   # The macro script
└─ README.md                  # This file

## 🧑‍💻 Author

Developed by [Abdulrahman M Hezam].  
Feel free to fork or contribute if you'd like to enhance it.

## 📄 License

This project is licensed under the MIT License.
