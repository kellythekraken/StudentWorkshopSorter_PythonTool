from openpyxl import load_workbook
import random, os, math

# TODO: Cleanup the workshop_nameiable name, e.g. consistent workshop OR class

# ============= Workshop variables to be customized =============================== #
excel_file_name = 'StudentWorkshop_SampleExcelSheet.xlsx'
wunsch_sheetname = 'Wuensche'
class_timetable_sheetname = 'Timetable'

last_name_column = 1 #A
first_name_column = 2 #B
name_start_row = 2 # the row where student's name start
workshop_name_start_column = 4 #D 
student_schedule_column = 15 #O

num_workshops_for_students = 5  # Number of workshops each student must take
num_workshop_rounds = 5   # Number of sessions/rounds each workshop provides

# ============== Do not edit the variables beyond this point! ===================== #
current_directory = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_directory, excel_file_name)
workbook = load_workbook(file_path)
sheet = workbook.get_sheet_by_name(wunsch_sheetname) 

# Load class name in order from excel
def load_workshop_names_from_excel():
    list = []
    current_column = workshop_name_start_column

    while True:
        workshop_name = sheet.cell(row = 1, column = current_column).value
        if workshop_name is not None:
            list.append(workshop_name)
            current_column += 1
        else:
            pref_end_column = current_column - 1
            break
    return list, pref_end_column

# Load student and group name from excel
def load_student_names_from_excel():
    names = []
    current_row = name_start_row
    
    while True:
        last_name = sheet.cell(row=current_row, column=last_name_column).value
        first_name = sheet.cell(row=current_row, column=first_name_column).value
        
        if last_name is not None and first_name is not None:
            full_name = f"{last_name} {first_name}"
            names.append(full_name)
            current_row += 1
        else:
            break
    return names

# Static workshop_nameiables
student_names = load_student_names_from_excel()
workshop_list, preference_end_column = load_workshop_names_from_excel()

# Erase the specified range of cells for student schedules
def Cells_Cleanup(start_row, start_col, end_row = None, end_col = None):
    #if end row or end col is not definied, the clean up will be executed until a blank row/col is met.
    if not end_row: 
        end_row = start_row + len(student_names)
    if not end_col:
        end_col = None

    for row in sheet.iter_rows(min_row=end_row + 1, min_col=4, max_col=4):
        if row[0].value is not None:
            end_row +=1
        else:
            break

    # Clean up cells from start_col to end_col until end_row
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for cell in row:
            cell.value = None

    workbook.save(excel_file_name)

# Create a dictionary with student name & their preference
# Also produce a list of student with no preference
def Fetch_Student_Preference_List():
    dict_student_pref = {}
    students_with_no_pref = []
    current_row = name_start_row
    workshop_choices = preference_end_column - workshop_name_start_column + 1 #calculate how many workshops to choose from

    # Loop through all students until the next empty row
    while True:
        last_name = sheet.cell(row=current_row, column=last_name_column).value
        first_name = sheet.cell(row=current_row, column=first_name_column).value
        
        # for each student:
        if last_name is not None and first_name is not None:
            full_name = f"{last_name} {first_name}"
            curr_pref_column = workshop_name_start_column
            int_preference_list = []

            # loop through all preference column and construct a list 
            for i in range(workshop_choices):
                current_pref = sheet.cell(row = current_row, column = curr_pref_column).value
                int_preference_list.append(current_pref)

                curr_pref_column += 1

                # NOTE: make sure to take into consideration some kids doesn't fill everything
                # CURRENTLY THE CODE WILL SKIP STUDENT THAT DOES NOT HAVE PREF FOR ALL WORKSHOP
                if current_pref is None:
                    int_preference_list = []
                    break
            
            # sort student and pref pair to dict ; student with no pref to list
            if not int_preference_list:
                students_with_no_pref.append(full_name)
            else:
                preference_list = Sort_Preference_List(int_preference_list)
                dict_student_pref[full_name] = preference_list

            current_row += 1
        else:
            break
    
    return dict_student_pref, students_with_no_pref

# Translate the value to class name and arrange them from favorite -> least favorite
def Sort_Preference_List(int_preflist):
    # int_preflist will looks like: [1,3,6,7,4,10]
    # sort the int_preflist to [1,2,3,4,5]
    temp_workshop_list = workshop_list
    zipped_lists = zip(int_preflist, temp_workshop_list)
    sorted_zipped_lists = sorted(zipped_lists, key=lambda x: x[0])
    sorted_workshop_preference = [item[1] for item in sorted_zipped_lists]
    #print(sorted_workshop_preference)

    # translate the number to class name [IT, Gastro Textil]
    return sorted_workshop_preference
student_preference_dict, students_with_no_preference = Fetch_Student_Preference_List()

# Calculate how many student should be in each class, if this num exceed 13 then keep it 13.
def Calculate_Max_Students_Per_Class():
    number_of_students = len(student_names)
    number_of_workshops = len(workshop_list)
    average = (number_of_students * num_workshops_for_students) / (number_of_workshops * num_workshop_rounds)
    average = math.ceil(average)
    if average > 13: average = 13
    return average

# Main logic to sort students to workshop based on their preference
def Sort_Student_to_Workshop_by_Preference():
    dict_student_classschedules = {key: [] for key in student_names}  # Keep track of each student and their class

    # Create dict consist of class name and 5 lists of student in each session
    class_dict = {} # {Textil:[[student A,B,C],[student D,E,F],[student G,H,I]]; }

    #keep track of the classes still available
    available_classes = workshop_list.copy()
    
    for a in range(len(workshop_list)):
        class_name = workshop_list[a]
        class_list = []
        for b in range(num_workshop_rounds):
            class_list.append([])
        class_dict[class_name] = class_list

    max_class_size = Calculate_Max_Students_Per_Class()

    # Assign classes according to student's preference
    for i in range(num_workshop_rounds):
        # loop through all students

        for student in student_preference_dict:
            # Get their 1st choice
            student_choice =  student_preference_dict[student][0]
            
            while student_choice not in available_classes:
                del student_preference_dict[student][0]
                student_choice =  student_preference_dict[student][0]

            # retrieve the list of sessions from the chosen class
            chosen_class_sessions = class_dict[student_choice]
            
            # Find the length of the smallest class in the available ones, sort student to that class
            session_with_least_students = min(chosen_class_sessions, key=len)

            # keep looping through the list of student's pref list until they're filled to the next class
            while len(session_with_least_students) >= max_class_size:
                # remove the preference for fully occupied class
                if student_choice in available_classes:
                    available_classes.remove(student_choice)
                    print("remove",student_choice,"from available class")
                
                del student_preference_dict[student][0]
                if not student_preference_dict[student]:
                    print("WARNING!",student,"cannot be assigned in ANY class because they're all full")
                    # do something!
                    break
                
                #get the next available class on the pref list
                student_choice = student_preference_dict[student][0]
                
                # loop and find the next available choice
                chosen_class_sessions = class_dict[student_choice]
                session_with_least_students = min(chosen_class_sessions, key=len)
            
            # remove the class from their list
            del student_preference_dict[student][0]
            # assign student to that session
            class_index = chosen_class_sessions.index(session_with_least_students)
            class_dict[student_choice][class_index].append(student)
            
            # put the class in the student's timetable
            # TODO: ALSO print the preference value? if possible, to check if it's fair.
            dict_student_classschedules[student].append(student_choice)

            # go to the next student

    print("available class left:", available_classes)

    # Assign classes for students with no preference
    for i in range(num_workshop_rounds):
        available_class_reset = False
        for student in students_with_no_preference:
            # get the list of class taken by current student, remove them from the available choice 
            classes_already_taken = dict_student_classschedules[student]
            class_to_choose_from = [x for x in available_classes if x not in classes_already_taken]

            # if there're no available class, open up class that are already filled.
            if not class_to_choose_from or not available_classes:
                available_classes = workshop_list.copy()
                available_class_reset = True
                class_to_choose_from = [x for x in available_classes if x not in classes_already_taken]
                print("No available classes found! Opening up the class number limitation.")

            # Create a temporary copy of the dict, remove the class that's already taken by student.
            new_class_dict = class_dict.copy()
            for key in class_to_choose_from:
                if key not in new_class_dict:
                    new_class_dict.pop(key)
            
            # Calculate the sum of total students for each workshop in the new dict
            total_student_in_each_workshop = {key: sum(len(sublist) for sublist in value) for key, value in new_class_dict.items()}

            # Pick the class with the least number of students in total
            class_with_least_students = min(total_student_in_each_workshop, key=total_student_in_each_workshop.get)
            chosen_workshop_sessions = class_dict[class_with_least_students]
            # take the session with least number of people
            session_with_least_students = min(chosen_workshop_sessions, key=len)

            if not available_class_reset:
                # if all sessions are full, remove this class from the available class
                while len(session_with_least_students) >= max_class_size:
                    #print("in the while loop")
                    if class_with_least_students in available_classes:
                        print(class_with_least_students,"is full!")
                        available_classes.remove(class_with_least_students)
                    
                    if not available_classes:
                        print("WARNING! No more available class!")
                        # available_classes = workshop_list
                        break
                    
                    random_class_choice = random.choice(class_to_choose_from)
                    chosen_workshop_sessions = class_dict[random_class_choice]
                    session_with_least_students = min(chosen_workshop_sessions, key=len)
            
            # Assign student to the session
            class_index = chosen_workshop_sessions.index(session_with_least_students)
            class_dict[class_with_least_students][class_index].append(student)
            
            # put the class in the student's timetable
            dict_student_classschedules[student].append(class_with_least_students)

    # Print the summary

    print("**********CLASS SUMMARY*****************")
    for key, value in class_dict.items():
        print(key)
        for sublist in value:
            print(len(sublist),":",sublist)

    print("\n\n**********STUDENT SCHEDULE*****************")
    for key, value in dict_student_classschedules.items():
        print(key)
        print(len(value),":",value)
   
    return dict_student_classschedules  

# Update excel with list of classes for each student 
def Excel_Update_Student_Schedule(student_schedule_dict):
    row = name_start_row # starting row 

    for student, classes in student_schedule_dict.items():
        col = student_schedule_column # Write each element from the list to columns D to H

        for _class in classes:
            sheet.cell(row=row, column=col, value = _class)
            col += 1
        
        row += 1  # Move to the next row for the next list
    workbook.save(file_path)

# Execution code
# Cells_Cleanup(2, 15, None, 19)

student_schedule = Sort_Student_to_Workshop_by_Preference()

Excel_Update_Student_Schedule(student_schedule)