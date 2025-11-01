# 🧾 Automated-Certificate-Generator
A Node.js app that auto-generates and converts DOCX templates with data into ready-to-use certificates and PDFs.

## 📖 Overview
This project allows you to upload any `.docx` template and generate certificates (or any similar document) dynamically using JSON data.  
It fills the placeholders in the template, merges all generated documents, and converts them into a single downloadable PDF file.

---

## 🚀 Features
- Accepts **any `.docx` template** with placeholders  
- Automatically fills placeholders with **JSON data**  
- Generates one certificate per record  
- Merges all certificates into a single `.docx`  
- Converts merged file into **PDF** automatically  
- Sends generated PDF as API response  

---

## 🛠️ Tech Stack
- **Node.js**  
- **Express.js**  
- **fs-extra**  
- **PizZip**  
- **Docxtemplater**  
- **DocxMerger**  
- **office-to-pdf**

---

## ⚙️ How It Works
1. Add your `.docx` template in the project folder.  
2. Send a POST request to the API endpoint with:
   - `templatePath` → path of the `.docx` file  
   - `Data` → array of objects containing placeholder values  
3. The system:
   - Fills the placeholders in the template  
   - Generates one document per data entry  
   - Merges all DOCX files  
   - Converts the merged DOCX into a single PDF  
4. The final PDF is returned in the response.

---

## 📩 Example Request (Postman)
**POST** `http://localhost:3000/api/generate`

**Body (JSON):**
```json
{
  "templatePath": "./templates/TransferCertificate.docx",
  "Data": [
    {
      "TC_NO": "191/2025/07",
      "NAME_OF_PUPIL": "AARAV SHARMA",
      "PEN": "9823456712",
      "MOTHER_NAME": "NEHA SHARMA",
      "FATHER_NAME": "RAHUL SHARMA",
      "NATIONALITY": "INDIAN",
      "CASTE": "GENERAL",
      "DOB": "12/08/2011",
      "DOB_IN_WORDS": "TWELFTH AUGUST TWO THOUSAND ELEVEN",
      "ISFAILED": "NO",
      "SUBJECTS": "ENGLISH, HINDI, MATHEMATICS, SCIENCE, SOCIAL STUDIES, COMPUTER, GENERAL KNOWLEDGE",
      "LAST_CLASS": "EIGHT",
      "LAST_RESULT": "QUALIFIED",
      "DUES_PAID": "YES",
      "FEE_CONCESSION": "NO",
      "NCC_SCOUT_GUIDE": "NO",
      "STRUCK_OFF_DATE": "31/03/2025",
      "REASON_FOR_LEAVING": "TRANSFER TO ANOTHER CITY",
      "TOTAL_MEETINGS": "215",
      "DAYS_ATTENDED": "198",
      "CONDUCT": "VERY GOOD",
      "SCHOOL_TYPE": "INDEPENDENT",
      "REMARKS": "REGULAR AND DISCIPLINED STUDENT",
      "ISSUE_DATE": "16/04/2025"
    },
    {
      "NAME_OF_PUPIL": "AARAV SHARMA",
      "PEN": "9823456712",
      "MOTHER_NAME": "NEHA SHARMA",
      "FATHER_NAME": "RAHUL SHARMA",
      "NATIONALITY": "INDIAN",
      "CASTE": "GENERAL",
      "DOB": "12/08/2011",
      "DOB_IN_WORDS": "TWELFTH AUGUST TWO THOUSAND ELEVEN",
      "ISFAILED": "NO",
      "SUBJECTS": "ENGLISH, HINDI, MATHEMATICS, SCIENCE, SOCIAL STUDIES, COMPUTER, GENERAL KNOWLEDGE",
      "LAST_CLASS": "EIGHT",
      "LAST_RESULT": "QUALIFIED",
      "DUES_PAID": "YES",
      "FEE_CONCESSION": "NO",
      "NCC_SCOUT_GUIDE": "NO",
      "STRUCK_OFF_DATE": "31/03/2025",
      "REASON_FOR_LEAVING": "TRANSFER TO ANOTHER CITY",
      "TOTAL_MEETINGS": "215",
      "DAYS_ATTENDED": "198",
      "CONDUCT": "VERY GOOD",
      "SCHOOL_TYPE": "INDEPENDENT",
      "REMARKS": "REGULAR AND DISCIPLINED STUDENT",
      "ISSUE_DATE": "16/04/2025"
    },
    {
      "NAME_OF_PUPIL": "AARAV SHARMA",
      "PEN": "9823456712",
      "MOTHER_NAME": "NEHA SHARMA",
      "FATHER_NAME": "RAHUL SHARMA",
      "NATIONALITY": "INDIAN",
      "CASTE": "GENERAL",
      "DOB": "12/08/2011",
      "DOB_IN_WORDS": "TWELFTH AUGUST TWO THOUSAND ELEVEN",
      "ISFAILED": "NO",
      "SUBJECTS": "ENGLISH, HINDI, MATHEMATICS, SCIENCE, SOCIAL STUDIES, COMPUTER, GENERAL KNOWLEDGE",
      "LAST_CLASS": "EIGHT",
      "LAST_RESULT": "QUALIFIED",
      "DUES_PAID": "YES",
      "FEE_CONCESSION": "NO",
      "NCC_SCOUT_GUIDE": "NO",
      "STRUCK_OFF_DATE": "31/03/2025",
      "REASON_FOR_LEAVING": "TRANSFER TO ANOTHER CITY",
      "TOTAL_MEETINGS": "215",
      "DAYS_ATTENDED": "198",
      "CONDUCT": "VERY GOOD",
      "SCHOOL_TYPE": "INDEPENDENT",
      "REMARKS": "REGULAR AND DISCIPLINED STUDENT",
      "ISSUE_DATE": "16/04/2025"
    }
  ]
}
