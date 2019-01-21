
import tkinter as tk
import tkinter.ttk as ttk
from PIL import ImageTk, Image
import tkinter.filedialog
from tkinter.filedialog import askopenfilename
from Student_tracker_V2 import *



img = Image.open(r"background.jpg")

class SampleApp(tk.Frame): #change to tk.Frame
    def __init__(self, parent, *args, **kwargs): #added parent
    
        tk.Frame.__init__(self, parent, *args, **kwargs) #added parent
        

        # Create all the frames for the GUI
        container = tk.Frame(parent) #changed self to parent
        self.bk_ground = ImageTk.PhotoImage(img) # creates the background image for all windows
        self.w = self.bk_ground.width() 
        self.h = 300
    
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        

        self.frames = {}
        for F in (StartPage, Finished_Page):
             # put all of the pages in the same location;
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

        
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.show_frame("StartPage")

    def show_frame(self, page_name):
        # function that switches between window
        frame = self.frames[page_name]
        frame.tkraise()
        

class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        
        # create canvas to link image and frame
        self.bk = tk.Canvas(self,width=self.controller.w, height=self.controller.h)
        self.bk.pack()
        self.bk.create_image(0, 0, image=self.controller.bk_ground, anchor='nw')
        
        # create labels, buttons in page
        self.canvas_title = self.bk.create_text(100, 30, anchor="nw", fill="Black", text="Adding your Kahoot Results ", font=("arial", 18))
        self.canvas_desciption = self.bk.create_text(125, 100, anchor="nw", fill="white", text="Search for your Kahoot File", font=("arial", 14))
        
        # creates the quit and continue button
        self.canvas_button_quit = tk.Button(self.bk,text = "Quit", command = root.destroy, anchor = "w")
        self.canvas_button_start = tk.Button(self.bk,text = "Find your File",command=lambda : self.open_file(controller), anchor = "w")
        self.button1_window = self.bk.create_window(225, 155, anchor="nw", window=self.canvas_button_quit)
        self.button2_window = self.bk.create_window(200, 125, anchor="nw",  window=self.canvas_button_start)
       
    def open_file(self,controller): 
        # gets the file address, display a message and a continue button
        fname = askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if fname:
            try:
                self.canvas_info = self.bk.create_text(65, 200, anchor="nw", fill="red", text="{} has loaded".format(fname), font=("arial", 8))
                self.canvas_button_contine = tk.Button(self.bk,text = "Merge Files",command= lambda : main(self,controller,fname) , anchor = "w")
                self.button3_window = self.bk.create_window(205, 220, anchor="nw", window=self.canvas_button_contine)
                                                           
            except:      
                self.canvas_info = self.bk.create_text(250, 100, anchor="nw", fill="red", text="Error -Failed to load", font=("arial", 14))
    
    
class Finished_Page(tk.Frame):
    # If displayed then results have been added
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
    
        # Create Canvas that links with image 
        self.bk1 = tk.Canvas(self,width=self.controller.w, height=self.controller.h)
        self.bk1.pack()
        self.bk1.create_image(0, 0, image=self.controller.bk_ground, anchor='nw')
        
        #Create Title
        self.canvas_title2 = self.bk1.create_text(45, 30, anchor="nw", fill="white", text="The Kahoot Results Have Been Added", font=("arial",18))
        self.canvas_button_quit_finshed = tk.Button(self.bk1,text = "Quit", command = root.destroy, anchor = "n")
        self.button1_window = self.bk1.create_window(225, 140, anchor="nw", window=self.canvas_button_quit_finshed)
        
    
        
class Error_Page():
    # all errors 
    
        
    def outlier(percentage,total_test,row_idx,column,name,correct_dic,total,current_pec_dic,read_student_df, total_points,controller):
        self = tk.Toplevel()
        '''This displays if a students result is very low. Asks user if they want to add results to the tracker '''
        
        # Display information and asks to add or get rid of results 
        label1 = ttk.Label(self, text="A Student has an unusually low score.", font=("arial",16))
        label1.pack()
        label2 = ttk.Label(self, text="Do you want to add it to the Tracker ?", font=("arial",12))
        label2.pack()
        label3 = ttk.Label(self, text="Student's Name :{}".format(name), font=("arial",12))
        label3.pack()
        label4 = ttk.Label(self, text="{}%".format(int(percentage)), font=("arial",12))
        label4.pack()
        # this will add the result to the tracker 
        B1 = ttk.Button(self,text = "YES", command =lambda: Error_Page.outlier_button(self,total_test,row_idx,column,name,correct_dic,total,current_pec_dic,
                                                                                                  read_student_df, total_points,'YES',percentage))
        B1.pack()
        
        # This will get rid of the results and won't update the student's percentage
        B2 = ttk.Button(self,text = "NO", command = lambda : Error_Page.outlier_button(self,total_test,row_idx,column,name,correct_dic,total,current_pec_dic,
                                                                                                 read_student_df, total_points,'NO',0))
        B2.pack()
        # makes the code wait for the this window to be destroyed before continuing 
        self.wait_window(self)
        
    def outlier_button(self,total_test,row_idx,column,name,correct_dic,total,
                       current_pec_dic,read_student_df, total_points,error,results):
        
        # Checks if this is the first test
        if total_test > 0 :
            # gets the current percentage that the student is at 
            current_percentage_object = read_student_df.cell(row_idx,  column)
            current_percentage = current_percentage_object.value
            
            # checks whether the user wants to update the student's percentage  
            if error == "NO":
                #does not update
                percentage = current_percentage
                    
            elif error == "YES":
                # updates the current percentage by adding the weighted value of this test to the current percentage
                correct = correct_dic[name]
                current_percentage = results / 100
                current_percentage = total_points * current_percentage 
                percentage = int(((current_percentage + correct ) / (total_points + total) * 100))
            
        # check whether this is the first value
        elif total_test == 0 :
            # checks whether the user wants to update the student's percentage  
            if error == "NO":
                # replace the percentage with nothing as the result should be null
                percentage =  " "

               
            elif error == "YES":
                # replace the percentage with the current result 
                percentage = results
            
        # updates the current percentage dictionary 
        current_pec_dic.setdefault(name, percentage)
             
        # destorys this window so that the process can start again     
        self.destroy()
        
    def place_student(name,student_list,result_students,result_df,student_questions,ignore_student):
        '''This displays if a student in the tracker can not be matched in the kahoot quiz  '''
        
        # creates the list for the drop down box
        finding_student = []
        for student in student_questions:
            finding_student.append(student[0])
        
        # adds a NO option to that list 
        finding_student.append("NO")
        
        self = tk.Toplevel()
        # displys information for the frame
        label1 = ttk.Label(self, text="Unmatched Student :{}".format(name), font=("arial",16))
        label1.pack()
        label2 = ttk.Label(self, text="Can you see the name they used ?", font=("arial",12))
        label2.pack()
        self.Combox = ttk.Combobox(self, values=finding_student, width=15)
        self.Combox.pack()
        B1 = ttk.Button(self,text = "YES", command =lambda: Error_Page.place_student_button(self,
                        self.Combox.get(),name,student_list,result_students,result_df,student_questions,ignore_student))
        B1.pack()
        # makes the code wait for the this window to be destroyed before continuing 
        self.wait_window(self)
            
            
    def place_student_button(self,name,new_name,student_list,result_students,result_df,student_questions,ignore_student):
        
        # if No then add name onto the ignore list for future use
        if name == "NO":
            ignore_student.append(new_name)
            
        else:
            # add the new name to result_student,result_df,student_question for future use
            result_students.append(new_name)
            result_df = result_df.replace(name,new_name)
            for student in student_questions:
                if student[0] == name:
                    student[0] = new_name
                    
        # destorys this window so that the process can start again   
        self.destroy()
        
        
    def find_column(title,column_names,error_column,tests_number_list,total_number_list,read_student_df):
        '''This displays if a column in the kahoot quiz can not be matched in the tracker  '''
        self = tk.Toplevel()
        
        # Displays the information
        label1 = ttk.Label(self, text="Unmatched Column :{}".format(title), font=("arial",16))
        label1.pack()
        label2 = ttk.Label(self, text="Which topic would you like to add the results to ?", font=("arial",12))
        label2.pack()
        self.Combox = ttk.Combobox(self, values=column_names, width=45)
        self.Combox.pack()
        B1 = ttk.Button(self,text = "YES", command =lambda: Error_Page.find_column_button(self,title, self.Combox.get(),column_names,error_column,
                                                                                          tests_number_list,total_number_list,read_student_df))
        B1.pack()
        # makes the code wait for the this window to be destroyed before continuing 
        self.wait_window(self)
            
            
    def find_column_button(self,title, column,column_names,error_column,tests_number_list,total_number_list,read_student_df):
        # Iterate through all existing columns in the tracker
        for counter,a_column in enumerate(column_names):
            # matches the newly select column 
            if column == a_column:
                # adds to the counter because of the layout of the tracker 
                counter += 2
                # gets the information needed from the tracker 
                tests_number_object = read_student_df.cell(4, counter) 
                tests_number = tests_number_object.value
                total_number_object = read_student_df.cell(3, counter) 
                total_number = total_number_object.value
                # add counter,total and test number to the list to be used in future 
                error_column.append(counter)
                total_number_list.append(total_number)
                tests_number_list.append(tests_number)
                
        # destorys this window so that the process can start again                
        self.destroy()
       
        
       
            
        
        
    
        
        
if __name__ == "__main__":
    root = tk.Tk() #Added
    app = SampleApp(root) #added root
    root.mainloop()

