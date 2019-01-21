# Student_Tracker

Aim 

The aim of the code is to allow users, primarily teachers, to  consolidate their student data quickly and efficiently. To accomplish this aim, I will create a program that :

Create a program that will load results from one excel sheet and combine it with data from another.

Creates a GUI for navigation purposes

Allows the user to select a file to merge with the tracker . 

Gets the results and updates the current percentage, total test done and total questions

Allows the user to not update outlier results that will distort the students results.

Allow the code to handle issues with matching students in the tracker and quiz by manually selecting the correct student.

Allow the code to handle issues with matching columns in the tracker and quiz by manually selecting the correct column.

Get the code to write the questions to that the student got wrong to the tracker. 

Design       

Program Structure

GUI :

  Start Page - Allows the user to select a file using a button and then continue with the program.
  
  Finish Page - Displays that the merging of the files have been completed. 
  
  Error Page  - Handles all error messages like outlier results and issues with matching students and topics. 

Backend

  Kahoot info -Gets all the information from the kahoot excel file such as questions asked ,results,  total questions ,title of the    quiz.

  Matching section -Matches the student in the kahoot excel file and the students in the tracker and matches the title in the kahoot file with the topic in the tracker file. If it is unable to match, the code will go to the error page. 

  Add percentage -Calculates the current percentage from the kahoot information. If other quizzes have already happened in the same topic, it will calculate the new current percentage. If the result is too low, it will go to the error page to ask the user whether they want to add the results to the tracker. 

  Add wrong questions -Insert all the questions that each student got wrong into the excel file.

Save file-Save the changes in the excel tracker file and load finish page. 
 

