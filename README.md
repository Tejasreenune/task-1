# Medical Appointment No-Show Dataset – Cleaning Project 🩺📊

🎯 Objective

To clean and prepare the raw “Medical Appointment No-Show” dataset by:

- Handling missing values

- Removing duplicates

- Fixing encoding issues in text (e.g., neighbourhood names)

- Standardizing date formats

- Renaming columns to consistent, clean names
- - Ensuring all data types are correct

🗂️ Dataset Info

- Source: [Kaggle Dataset] (https://www.kaggle.com/datasets/joniarroba/noshowappointments)

- Size: ~110K records

- Fields: PatientID, AppointmentDay, Gender, Neighbourhood, No-show, etc.

🧹 Key Cleaning Steps

- Cleaned corrupted neighbourhood encodings (e.g., "Ã‰" → "É")

- Converted "ScheduledDay" and "AppointmentDay" to proper Excel date format

- Created additional fields: "Wait Days", "Appointment Date", etc.

- Standardized columns to "snake_case" format
  
🔣 Excel Formulas Used for Data Cleaning
Below are the formulas i have applied to clean and transform the dataset using Excel:

| Column Name               | Description                                                  | Formula                                             |
|---------------------------|--------------------------------------------------------------|-----------------------------------------------------|
| 'ScheduleDay_Cleaned'     | Converted 'ScheduledDay' from datetime to 'dd-mm-yyyy'       | '=DATEVALUE(LEFT(E2,10))'                           |
| 'AppointmentDay_Cleaned'  | Converted 'AppointmentDay` to 'dd-mm-yyyy' format            | '=DATEVALUE(LEFT(F3,10))'                           |
| 'Wait_Days'               | Calculated number of days between scheduled and appointment  | '=H2 - G2'                                          |
| 'Gender_Cleaned'          | Converted Gender (F, M) value to (Female, Male)              | '=IF(C2="F", "Female", IF(C2="M", "Male", ""))'                                          |
| 'Age_Groups'              | Categorized patients into age groups                         | '=IF(J2<18,"Child",IF(J2<=60,"Adult","Senior"))'    |
| 'Neighbourhood_Cleaned'   | Standardized neighborhood names to proper case               | '=PROPER(L2)'                                       |
| 'No_Show_Cleaned'         | Converted `No_show` to binary (0=No, 1=Yes)                  | '=IF(T2="Yes", 1, 0)'                               |


📁 Final Files

- "KaggleV2-May-2016.csv" – Raw dataset

- "medical-no-show-cleaning.xlsx" – cleaned dataset

🛠️ Tools Used

- Microsoft Excel

👀 Preview
<img width="1366" height="322" alt="task-1" src="https://github.com/user-attachments/assets/a65bcbdc-56ea-4eca-8035-4840fd85b983" />

**✍️ Author**

**NUNE TEJASREE**
