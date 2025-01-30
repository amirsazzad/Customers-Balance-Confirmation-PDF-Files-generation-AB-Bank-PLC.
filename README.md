# Customers Balance Confirmation (PDF) Files generation AB-Bank-PLC.
Class Name: PDFGen  Purpose: This class is responsible for generating PDF documents containing balance confirmation details for bank accounts. The PDFs are password-protected and stored in a specified directory. The class also logs the generated PDFs into a database for tracking purposes.
PDF Generation and Email Notification System
This project is a C# application designed to generate PDF documents containing balance confirmation details for bank accounts. The generated PDFs are password-protected and stored in a specified directory. The application also logs the generated PDF details into a database for tracking purposes.

<h1 align="center">📄 PDF Generation System</h1>

<p align="center">
  <strong>A C# application for generating password-protected PDFs and logging details into a database.</strong>
</p>

---

## 🚀 **Features**
- **PDF Generation**: Dynamically generates PDFs with account details, balance confirmation, watermarks, and logos.
- **Password Protection**: Secures each PDF with a password derived from the account number.
- **Database Integration**: Fetches account details and logs PDF information into a database.
- **Configuration Management**: Uses `App.config` for settings like database connections and file paths.
- **Error Handling**: Robust error handling and logging for smooth execution.

---

## 🛠️ **Prerequisites**
Before running the application, ensure you have the following installed:
- **.NET Framework** (version 4.5 or higher)
- **iTextSharp** library (for PDF generation)
- **SQL Server** or any compatible database
- **IBMDA400.DataSource.1** provider (for database connectivity)
- 

## 🚦 **Setup Instructions**

### 1. **Clone the Repository**
Clone this repository to your local machine:
bash
Copy
git clone https://github.com/amirsazzad/PDF-Generator-For-AB-Bank-PLC..git

2. Configure the Application
Open the App.config file and update the following settings:
Database Connection Strings:
xml
Copy
<connectionStrings>
  <add name="ConnectionString1" connectionString="Your_Database_Connection_String" />
</connectionStrings>
Run HTML
Application Settings:

xml
Copy
<appSettings>
  <add key="Type" value="Your_Account_Type" />
  <add key="Exclude" value="Excluded_Branches" />
  <add key="CBS" value="Your_CBS_Value" />
  <add key="Unit" value="Your_Unit_Value" />
  <add key="pathpdf" value="Path_To_Save_PDFs" />
  <add key="relativePath" value="Path_To_Watermark_Image" />
  <add key="imageWM_path" value="Path_To_Image_Folder" />
  <add key="relativeURL" value="Path_To_Logo_Image" />
  <add key="imageURL_path" value="Path_To_Logo_Folder" />
</appSettings>
Run HTML
Ensure the database table pdfrange exists with the following schema:

sql
Copy
CREATE TABLE pdfrange (
    cntfrom INT,
    cntto INT
);
Ensure the database table Attachment exists with the following schema:
sql
Copy
CREATE TABLE Attachment (
    Branch NVARCHAR(255),
    Internal_AC NVARCHAR(255),
    External_AC NVARCHAR(255),
    Pdffile NVARCHAR(255),
    Email NVARCHAR(255),
    sms NVARCHAR(MAX),
    Process_date DATETIME,
    issent BIT,
    seq NVARCHAR(255),
    SL INT
);
3. Install Dependencies
Install the iTextSharp library via NuGet Package Manager:

bash
Copy
Install-Package itextsharp
4. Build and Run
Open the solution in Visual Studio.

Build the project to resolve dependencies.

Run the application.

Usage
Input Credentials:
Provide the cbs_username and cbs_password when prompted to authenticate with the database.
PDF Generation:
The application will fetch account details from the database, generate PDFs, and save them to the specified directory.
Database Logging:
Details of the generated PDFs (e.g., file path, email, SMS) will be logged into the Attachment table.
Error Handling:
Any errors encountered during execution will be logged and displayed in the console.

File Structure
BALCON_APP/
├── App.config                # Configuration file for database and application settings
├── PDFGen.cs                 # Main class for PDF generation and database operations
├── README.md                 # This file
├── Image/                    # Folder containing watermark and logo images
│   ├── WMlogo.png            # Watermark image
│   └── ablogo.png            # Logo image
└── PDFs/                     # Folder where generated PDFs are saved

Example Output
Generated PDF:
Contains account details, balance confirmation, and a watermark.
Saved in the PDFs folder with a filename like BC1234-231015-000001.pdf.

Database Log:
A new entry is added to the Attachment table with details of the generated PDF.

<h4>Troubleshooting</h4>
No Rows Found:
Ensure the pdfrange table contains valid cntfrom and cntto values.
Verify that the BALCON table in the database contains the required account details.
Connection Errors:
Double-check the connection strings in App.config.
Ensure the database server is running and accessible.

PDF Generation Errors:
Verify that the image paths in App.config are correct.

After generating the PDF, an email can be sent.

Ensure the PDFs folder has write permissions.
Contributing
Contributions are welcome! If you find any issues or have suggestions for improvement, please open an issue or submit a pull request.
