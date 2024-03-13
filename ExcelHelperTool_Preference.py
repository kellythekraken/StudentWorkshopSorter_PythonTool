from openpyxl import load_workbook
import random, os

# COLUMN: VERTICAL NEWSPAPER COLUMNS | ROW: HORIZONTAL SEATS
# NOTE: Make the length of the column read automatically
# TODO: Cleanup the variable name, e.g. consistent workshop OR class

# Modifiable Variables
excel_file_name = 'StudentWorkshop_SampleExcelSheet.xlsx'
wunsch_sheetname = 'Wuensche'
class_timetable_sheetname = 'Timetable'

last_name_column = 1 #A
first_name_column = 2 #B
name_start_row = 2 # the row where student's name start
preference_start_column = 4 #D
preference_end_column = 13 #M

num_workshops_for_students = 5  # Number of workshops each student must take
num_workshop_rounds = 5   # Number of rounds of each workshop class

current_directory = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_directory, excel_file_name)
workbook = load_workbook(file_path)
sheet = workbook.get_sheet_by_name(wunsch_sheetname) 

# Load class name in order from excel
def load_workshop_names_from_excel():
    list = []
    num = preference_end_column - preference_start_column + 1
    current_column = preference_start_column
    for i in range(num):
        var = sheet.cell(row = 1, column = current_column).value
        list.append(var)
        current_column += 1
    return list

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

# Static variables
workshop_list = load_workshop_names_from_excel() 
student_names = load_student_names_from_excel()

# Erase the specified range of cells for student schedules
def Cells_Cleanup(start_row, start_col, end_col):
    end_row = start_row + len(student_names)

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
    workshop_choices = preference_end_column - preference_start_column + 1 #calculate how many workshops to choose from

    # Loop through all students until the next empty row
    while True:
        last_name = sheet.cell(row=current_row, column=last_name_column).value
        first_name = sheet.cell(row=current_row, column=first_name_column).value
        
        # for each student:
        if last_name is not None and first_name is not None:
            full_name = f"{last_name} {first_name}"
            curr_pref_column = preference_start_column
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

# NOTE: (statistic data can be collected at this point)

# Calculate how many student should be in each class, if this num exceed 13 then keep it 13.
def Calculate_Max_Students_Per_Class():
    number_of_students = len(student_names)
    number_of_workshops = len(workshop_list)
    average = (number_of_students * num_workshops_for_students) / (number_of_workshops * num_workshop_rounds)
    average = round(average) - 1
    if(average > 13): average = 13
    return average

# Main logic to sort students to workshop based on their preference
def Sort_Student_to_Workshop_by_Preference():
    dict_student_classschedules = {key: [] for key in student_names}  # Keep track of each student and their class

    # Create dict consist of class name and 5 lists of student in each session
    class_dict = {} # {Textil:[[student A,B,C],[student D,E,F],[student G,H,I]]; }

    #keep track of the classes still available
    available_classes = workshop_list
    
    for a in range(len(workshop_list)):
        class_name = workshop_list[a]
        class_list = []
        for b in range(num_workshop_rounds):
            class_list.append([])
        class_dict[class_name] = class_list

    max_class_size = Calculate_Max_Students_Per_Class()

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

    # go through the list of student with no pref in the end
    
    for i in range(num_workshop_rounds):
        for student in students_with_no_preference:
            # get all available class
            random_class_choice = random.choice(available_classes)
            # pick a random class, access their list of 5 sessions

            # take the session with least number of people
            # if all sessions are full, remove this class from the available class
            # if there're no available class, fill the student to a random class?

            smallest_class_length = min(len(lst) for lst in dict_classes_count.values())

            # Get all keys with lists of that length
            smallest_classes = [key for key, lst in dict_classes_count.items() if len(lst) == smallest_class_length]

            # available_class_for_student
            available_class_for_student = [idx for idx in smallest_classes if idx not in dict_students_assigned_classes.get(name, [])]

            if not available_class_for_student: #if all are taken, use the original complete list
                available_class_for_student = [idx for idx in class_names if idx not in dict_students_assigned_classes.get(name, [])]
            
            chosen_class = random.choice(available_class_for_student) 

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
    
'''
    for round in range(1, num_workshops + 1): 
        
        dict_classes_count = {key: [] for key in workshop_list} #dict_student_pref with {classes : [student1,2,3]}

        for name in student_names:
            # Find the length of the smallest class in the available ones
            smallest_class_length = min(len(lst) for lst in dict_classes_count.values())

            # Get all keys with lists of that length
            smallest_classes = [key for key, lst in dict_classes_count.items() if len(lst) == smallest_class_length]

            # available_class_for_student
            available_class_for_student = [idx for idx in smallest_classes if idx not in dict_student_classschedules.get(name, [])]

            if not available_class_for_student: #if all are taken, use the original complete list
                available_class_for_student = [idx for idx in class_names if idx not in dict_student_classschedules.get(name, [])]
            
            chosen_class = random.choice(available_class_for_student) 

            # Add the new class for this student to the dictW
            dict_student_classschedules[name].append(chosen_class)
            # add this student to dict_classes_count to keep track how many are in each every round
            dict_classes_count[chosen_class].append(name)

        #after looping through all students, write down dict_classes_count before they're renewed
#        Excel_Update_Students_In_Class(dict_classes_count,round)
'''

student_schedule = Sort_Student_to_Workshop_by_Preference()
