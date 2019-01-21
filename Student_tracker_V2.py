import pandas as pd 
import xlrd
from xlutils.copy import copy
import re

'''Function Section'''

def add_percentage(student_list,column,percentage_dic,total_points,total_test,total,new_write_tracker,ignore_student):
    '''writes the percentage and increments the total points and total tests'''
    new_write_student_df = new_write_tracker.get_sheet('Detail')# Opens new sheet because code had to save and close old sheet 
    new_write_student_df.write(4,column,total_test + 1) # Writes the 1 increment to total tests done. 
    new_write_student_df.write(3,column,total_points+total ) # Adds points to the total points of all the tests  
    row = 8
    
    #iterates through all the students
    for name in student_list:
        # stops if a student is in the ignore list
        if name not in ignore_student:
            value = percentage_dic[name] # finds current percentage from dictionary 
            new_write_student_df.write(row,column,percentage_dic[name]) # writes percentage to excel  
            row +=1
        else :
            row+=1
   
   
        
        
        
def add_wrong_answer(student_list,title,student_questions,questions,new_tracker_df,new_write_tracker,
                     current_pec_dic,ignore_student):
    '''writes the questions that each student got wrong into the tracker '''
    # iterates through the quiz student list - this is a list of lists where it contains the student name and all wrong answers
    for student in student_questions:
        student_name = student[0]
        # confirms that the student matches to the tracker student list
        if student_name in student_list :
            # checks to see if the student chould be ignored 
            if current_pec_dic[student_name]!= " ":
                # checks to see if the student has any wrong asnwers.
                if len(student) > 1 :
                    #get the sheet to be able to write to 
                    read_student_topics = new_tracker_df.sheet_by_name(student_name)
                    write_student_incorrect_question= new_write_tracker.get_sheet(student_name)
                    # set flags
                    counter = 0 
                    search = True
                    while search == True:
                        # finds the correct column to add the information to
                        column_info = read_student_topics.cell(0,counter)
                        column_title = column_info.value
                        if title == column_title.upper() :
                            #set flags
                            column = counter
                            row = 0 
                            adding_row = True
                            while adding_row ==True:
                                # finds an empty cell to write the code to
                                try :
                                    cell_info = read_student_topics.cell(row,column).value
                                    row+=1
                                except :
                                    # iterate through the wrong answers and writes them to the tracker line by line
                                    this_row = row
                                    for question in student[1:]:
                                        index = int(question)
                                        write_student_incorrect_question.write(this_row,column,questions[index - 1])
                                        adding_row = False
                                        search = False
                                        this_row += 1
                                row += 1
                        counter+=1
    

                            
def create_sheets(student_list,title,write_tracker,read_student_df):
    '''Creates the sheets in the tracker '''
    # iterates through the students
    for name in student_list:
        all_columns = []
        try : 
            # create the sheet with the student's name
            add_sheet = write_tracker.add_sheet(name)
            write_single_student_df = write_tracker.get_sheet(name)
            counter = 2 
            while True :
                # gets all the topic columns
                try : 
                    column_info = read_student_df.cell(7, counter)
                    column_title  = column_info.value
                    all_columns.append(column_title)
                    counter += 1
                except :
                        break
        except:
                print("Error in creating a new student's page")
            
        # writes all the topic columns in the new student's sheet
        for index in range(0,len(all_columns)):
                write_single_student_df.write(0,index,all_columns[index])
    
    #Saves the excel file        
    write_tracker.save('student_tracker.xls')
    

def match_student_names(name, student_questions,result_students,result_df,second_name,counter,student_list,ignore_student):
    '''Matches or change the names for student_questions and result_students to the excel tracker'''
    error = True
    match_counter = 0
    #Iterates thought the students
    for student in student_questions :
        student_name = student[0]
        # if there is a match change all information to the correct student name
        if name in student_name.upper() or second_name == student_name.upper():
            student[0] = name
            result_students[counter]= name
            result_df.iat[match_counter, 1] = name
            error = False
        match_counter += 1
    
    # if there is an error go to the error page to allow user to select a student to match the results too
    if error == True:
        from GUI import Error_Page
        Error_Page.place_student(name,student_list,result_students,result_df,student_questions,ignore_student)
        

def info(file):
    ''' Get info from the Kahoot file '''
    result_df = pd.read_excel(file, header = None, sheet_name="Final Scores")
    question_df = pd.read_excel(file, skiprows = 2 ,sheet_name="Question Summary")
    total_df = pd.read_excel(file, skiprows = 2 ,sheet_name="Overview")

    # Gets all the questions from the kahoot file 
    question_df = question_df.set_index('Players')
    question_df = question_df[pd.notnull(question_df.index)]
    

    all_question_columns = question_df.columns
    answer_columns = get_columns_from_question_df(all_question_columns)
    answer_df = question_df[answer_columns]
    

    
    student_questions = get_questions(answer_df)
    question_df = question_df.drop(answer_columns,1)
    questions = question_df.columns
    questions = questions[2:]
    
    #  Cleans the result dataframe to be used 
    title = result_df.iloc[0,0]
    title = topic_reduction(title)

    result_df = result_df.rename(columns={1 : 'Student', 2: 'Points', 3: "Correct", 4: "Incorrect"})
    result_df = result_df.iloc[3:,1:]
    result_df = result_df.dropna()
    result_df["Total"] = result_df.iloc[:,2] + result_df.iloc[:,3]
    result_df["Student"] = result_df["Student"].apply(lambda x : x.upper())
    result_df = result_df.reset_index()
    
    
    #gets the total to be used 
    total = get_total(total_df.iloc[1,1])
    
    return questions, result_df, total,title,student_questions


def get_correct_answers(name,result_students,counter,result_df,correct_dic,incorrect_dic,student_list):
    ''' Saves the current questions and incorrect into the dictionaries to be used to get current percentage '''
    # iterates throught the list of lists in result_students and check if the first or second name matches the name from the quiz
    if any(name in s for s in result_students):
        try:
            # gets information using name
            correct = int(result_df.Correct[result_df['Student'] == name])
            incorrect = int(result_df.Incorrect[result_df['Student'] == name])
            
        except:
            # gets information using counter 
            correct = int(result_df.Correct[result_df['Student'] == result_students[counter] ])
            incorrect = int(result_df.Incorrect[result_df['Student'] == result_students[counter] ])
        correct_dic.setdefault(name, correct)
        incorrect_dic.setdefault(name, incorrect)
        
    else:
        print(name)
       
        
    counter += 1
    return counter


        

def get_current_percentage(total_test,row_idx,column,name,correct_dic,total,
                           current_pec_dic,read_student_df, total_points ,controller,result_students,ignore_student):
    ''' Saves the current questions and incorrect into the dictionaries to be used to get current percentage '''
    
    # checks if name should be ignored
    if name not in ignore_student:
        correct = correct_dic[name]
        percentage = (correct / total) * 100
   
    # checks to see if result is an outlier 
        if percentage < 30 :
            from GUI import Error_Page
            Error_Page.outlier(percentage,total_test,row_idx,column,name,correct_dic,total,current_pec_dic,read_student_df, total_points,controller)
        
        else : 
            # checks to see if this is the first test
            if total_test > 0 :
                current_percentage_object = read_student_df.cell(row_idx,  column)
                current_percentage = current_percentage_object.value
                # combines current percentage with previous scores
                try : 
                    current_percentage = int(current_percentage / 100)
                    current_percentage = total_points * current_percentage 
                    percentage = ((current_percentage + correct ) / (total_points + total) * 100)
                except TypeError:
                    pass
                except:
                    print("Error")
            # if it is the first test, use the current percentage     
            elif total_test == 0 :
                percentage = percentage
        
    
        current_pec_dic.setdefault(name, int(percentage))
    


def get_columns_from_question_df(columns):
    ''' Normalise the column data from the kahoot quiz'''
    Q_columns = [columns[2]]
    for column in columns:
        if re.match(r"Q[1-9]+", column):
            Q_columns.append(column)
    return Q_columns


def get_questions(df):
    ''' iterate throught dataframe and find out all questions that the student got wrong'''
    all_questions = []
    # iterate throught dataframe
    for index, row in df.iterrows():
        equal_columns = []
        values = list(row)  # Get the unique values in the row
        equal_columns.append(index) # make the first value a name
        for counter, v in enumerate(values): # itertae through values
            if v == 0 : # if value is wrong append it to the equal_column list
                columns = df.columns[counter-1]
                equal_columns.append(columns[1:])
              
        all_questions.append(equal_columns) # make list of lists of all wrong questions 
    return all_questions

def get_total(value):
     ''' Finds out the total of the kahoot quiz'''
    new_value = value.split()
    total = new_value[0]
    total = int(total)
    return total

def get_total_test_num_colum(title,num_columns,read_student_df):
    ''' Finds out the tests_number,total_number,column,title'''
    error_column = []
    found = False
    column_names = []
    tests_number_list =[]
    total_number_list = []
    column_names = []
    for col in range(2,num_columns):
        column_info_object = read_student_df.cell(7, col)
        column_info =column_info_object.value # Get cell Object
        
        if column_info.upper() == title.upper(): # Compares topic with quiz title
            tests_number_object = read_student_df.cell(4, col) 
            tests_number = tests_number_object.value
            total_number_object = read_student_df.cell(3, col) 
            total_number = total_number_object.value
            found = True
            column = col
            column_names.append(column_info.upper())
            
        else:
            column_names.append(column_info.upper())
            
        
    if found == False :
        from GUI import Error_Page
        Error_Page.find_column(title,column_names,error_column,tests_number_list,total_number_list,read_student_df)
           
    if len(error_column) == 1 :
        column = error_column[0]
        tests_number =tests_number_list[0]
        total_number= total_number_list[0]
        title = column_names[column-2]
        
    
    return tests_number,total_number,column,title


def topic_reduction(topic):
    ''' Normalises the title'''
    result = ''.join([i for i in topic if not i.isdigit()])
    return result


def main(self,controller,file):
    from GUI import SampleApp
    '''Main code'''
    
    
    #get kahoot info
    questions,result_df,total,title,student_questions = info(file) 


    #Opens Excel file and allow to read and write to main sheet
    tracker_df = xlrd.open_workbook('student_tracker.xls') 
    read_student_df =tracker_df.sheet_by_name('Detail')
    write_tracker = copy(tracker_df)
    write_student_df = write_tracker.get_sheet('Detail')
    all_sheet_names = tracker_df.sheet_names()   
    


    #key variables 
    result_students = list(result_df["Student"]) # Turns results student name into a list
    correct_dic = {}
    incorrect_dic = {}
    current_pec_dic = {}
    student_list = []
    ignore_student = []
    num_columns = read_student_df.ncols
    total_test,total_points,column,title = get_total_test_num_colum(title,num_columns,read_student_df) # gets total point and test and column that we will write in 
    

    counter = 0 # counter needed for dict info 
    for row_idx in range(8, read_student_df.nrows):# get all the information
        # Collect correct and incorrect data
        cell_info = read_student_df.cell(row_idx, 0)  # Get cell object by row, col
        name = cell_info.value  # Gets Students name from tracker
        name = name.upper() # Makes it upper case
        cell_info_second = read_student_df.cell(row_idx, 1)  # Get cell object by row, col
        second_name = cell_info_second.value
        second_name = second_name.upper()
        student_list.append(name)
        match_student_names(name, student_questions,result_students,result_df,second_name,
                            counter,student_list,ignore_student)
        counter= get_correct_answers(name,result_students,counter,result_df,correct_dic,incorrect_dic,student_list)
        get_current_percentage(total_test,row_idx,column,name,correct_dic,total,
                               current_pec_dic,read_student_df, total_points ,controller,result_students,ignore_student)
    
    
    
    new_student_list = [] # list to new students that need to be added to the sheets
    add_new_students = False
    for student in student_questions:
        student_name = student[0]
        if student_name not in all_sheet_names and student_name not in ignore_student:
            new_student_list.append(student_name)
            add_new_students = True
        
    
   
    new_worksheet = False      
    if add_new_students == True:
        # create sheets for students if needed,saves it and loads the updated file
        new_worksheet = create_sheets(student_list,title,write_tracker,read_student_df)
        new_worksheet = True
    
    
    if new_worksheet == True :
        tracker_df .release_resources()
        del tracker_df 
        new_tracker_df = xlrd.open_workbook('student_tracker.xls')
        new_write_tracker = copy(new_tracker_df)
    else : 
        new_write_tracker = write_tracker
        new_tracker_df = tracker_df
        
    
        
    # adds percentage to sheet   
    add_percentage(student_list,column,current_pec_dic,total_points,total_test,total,new_write_tracker,ignore_student)


    #adds wrong questions to added sheets 
    add_wrong_answer(student_list,title,student_questions,questions,new_tracker_df,new_write_tracker,
                     current_pec_dic, ignore_student)

   
    # Saves the percentages in detail sheet and wrong questions
    new_write_tracker.save('student_tracker.xls')
 
    SampleApp.show_frame(controller,'Finished_Page')
    
    
