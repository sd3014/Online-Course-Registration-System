# Online-Course-Registration-System
An efficient Online Course Registration System using Python and Tkinter for student enrollment. It validates inputs and stores data in Backend_data.xlsx. Users can enter details, submit, clear, or exit the form. Built with Tkinter, OpenPyXL, Regex, and PIL, ensuring a smooth and error-free registration process. 🚀

### **Online Course Registration System** 

The **Online Course Registration System** is a Tkinter-based Python application that allows students to register for courses efficiently. It provides a user-friendly interface with proper validation and stores data in an Excel file.  

#### **Key Features:**  
✅ **User-Friendly Interface** – The system includes an intuitive GUI with labeled input fields and dropdown menus for gender, course selection, and duration. A welcome screen provides easy navigation.  

✅ **Student Registration Form** – Users can enter their **name, contact number, age, gender, address, email, course, and duration**. Validation ensures correct input before submission.  

✅ **Data Validation** –  
- **Name Check:** Allows only characters (max 32).  
- **Phone Number:** Validates 10-digit numeric input.  
- **Email Format:** Ensures a correct email structure using regex.  
- **Mandatory Fields:** Ensures all fields are filled before submission.  

✅ **Data Storage in Excel** –  
- The system stores user details in **Backend_data.xlsx** using **OpenPyXL**.  
- If the file does not exist, it is created automatically.  
- Entries are dynamically added to the next available row.  

✅ **Recent Data Confirmation** – After registration, the latest entry is retrieved and displayed as a confirmation message.  

✅ **Form Operations** –  
- **Submit Button:** Saves data and confirms registration.  
- **Clear Button:** Resets all fields for a new entry.  
- **Exit Button:** Closes the registration window.  

✅ **Graphical Enhancements** – The system includes background images and icons for a visually appealing interface.  

#### **Technologies Used:**  
🖥 **Python (Tkinter)** – GUI Design  
📂 **OpenPyXL** – Excel Data Management  
🔍 **Regex** – Input Validation  
🖼 **Pillow (PIL)** – Image Handling  

This **Online Course Registration System** ensures smooth and accurate student enrollments with a simple and interactive user experience. 🚀
