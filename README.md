# Grade-Wizard
Written in VBA utilizing MS excel macros, UserForms, and SQL queries, users can extract enrolment information and course grade reports from MS Access databases containing academic data. The grade report will be saved to the user's PC in an MS Word document. Data is imported and read from the database using SQL queries.

*This application was originally created for course credit in "Windows Application Programming"(CP 212) at Wilfrid Laurier Univeristy, taught by Dr. Hieder Ali (Fall 2022).*

Author: Erin Manson

To use the application you must have Registrar.accdb(or file in same format), and GradeWizardMaster.xlsm(or a .xlsm file with all modules and userforms imported) in the same directory.

Description:
This student grade application is run from the excel file GradeWizardMaster.xlsm. This file imports a database containing student and course information in the same format of Registrar.accdb. The user can select to receive a course enrolment summary on a new sheet OR a word document report detailing the final grades for a course. 

Step by step walkthrough:
1.	After opening the excel file you will see instructions and a button located on “Home Instruction Sheet”, press the button to launch the first UserForm in step one of the application.
(The button is attached to the sub “MainS1” in “ModuleStep1” this sub will launch the UserForm: “UserFormA5Step1”.)
 
2.	The userForm will make you select either “Generate word report on final grades for course” OR “Display course enrolment for course”. If no option is selected and the user presses continue a message box will show explaining the issue. Press cancel or X to close, and Continue to proceed once ready.
  
3.	Regardless of the user choice a FileDialog Box will be launched for the user to select a database to import course data from. This is done by either MainReportStep2 or MainCourseInfoStep2 subroutine dependant on the user selection in UserFormA5Step1.
(The FileDialog Box will only allow Access database files to be selected and only one file of this type may be selected at a time.)
 
4.	After an Access database is selected the second UserForm will be launched: UserFormA5Step2. This UserForm will populate a list box with course codes from the database. At this stage if the database does not contain the correct course table in the right format Error Handling will be utilized, and a message will be shown to the user.
 

5.
 A) If the user selected Course enrolment information on new sheet:
  A new sheet will be populated via the sub “GetCourseInfo()” contained in Module “Module2C”. The new sheet will be titled “*Course Name* Enrolment” and will contain student first, last names and student IDs who are enrolled in the course selected(See below).
  
<img width="259" alt="image" src="https://user-images.githubusercontent.com/126124271/220813212-d3426377-7f0f-459e-977d-d6e5995beaae.png">

 B) If the user selected Report of final grades on word document:
  A new word document in the user’s Documents folder will be saved with the name “*Course Name* Final Grade Report”. The word document report contains a Histogram chart of the final grades for the course. Final grades are calculated with the weight: Assignments 5%, Midterm 30%, and Final 50%, of the weights you wish to use are different you must alter the code in GetGradeReport() sub. All assignment averages are imported from the access database for each student. Following the chart the document contains other course statistics: average, min, max, mode, median, and standard deviation. 
 After creating a report, the message box will show that the program has finished and guide you to the report in your documents folder.

Example of Word document report that the application can produce for course "PC120" screenchot .png:
<img width="328" alt="image" src="https://user-images.githubusercontent.com/126124271/220812647-d8681341-28b4-4b10-aa10-1cb191d20f27.png">

 
The Application is now complete! The student marking/grade wizard application effectively handles errors, imports data from a database using SQL Queries, and makes a new sheet/word document dependant on UserForm input.


Email me at 0404em@gmail.com if you have any questions!

EEM

