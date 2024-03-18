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

def sublist_length(sublist):
    return len(sublist) if sublist is not None else float('inf')  # Use infinity for None to ensure they are placed at the end

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

# Return bool to check if all sessions of a class is full
def Remove_Full_Class(max_class_size, dict,available_classes):
    max_class_size_total = max_class_size * num_workshop_rounds
    for key,list_of_sublists in dict.items():
        total_students = sum(len(sublist) for sublist in list_of_sublists)
        if total_students >= max_class_size_total:
            if key in available_classes:
                print(key,"is now completely full!")
                available_classes.remove(key)

# Main logic to sort students to workshop based on their preference
def Sort_Student_to_Workshop_by_Preference():
    dict_student_classschedules = {key: {} for key in student_names}  # Keep track of each student and their class

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
            '''
            while student_choice not in available_classes:
                del student_preference_dict[student][0]
                student_choice =  student_preference_dict[student][0]
            '''
            # retrieve the list of sessions from the chosen class
            chosen_class_sessions = class_dict[student_choice]
            
            # Filter out the index of session that student's are already occupied with
            classindexes = list(dict_student_classschedules[student].values())
            classindexes.sort()
            #the classindexes calculation is correct, but somehow still taken into account and not excluded.

            available_sessions_to_take = chosen_class_sessions.copy()
            for i in classindexes:
                available_sessions_to_take[i] = None

            # Find the length of the smallest class in the available ones, sort student to that class
            session_with_least_students = min(filter(lambda x: x is not None, available_sessions_to_take), key=sublist_length)

            # keep looping through the list of student's pref list until they're filled to the next class
            while len(session_with_least_students) >= max_class_size:
                # remove the preference for occupied class
                del student_preference_dict[student][0]

                if not student_preference_dict[student]:
                    print("WARNING!",student,"cannot be assigned in ANY class because they're all full.")
                    # do something!
                    break
                #get the next available class on the pref list
                student_choice = student_preference_dict[student][0]
                
                chosen_class_sessions = class_dict[student_choice]
                available_sessions_to_take = chosen_class_sessions.copy()
                for index in classindexes:
                    available_sessions_to_take[index] = None

                session_with_least_students = min(filter(lambda x: x is not None, available_sessions_to_take), key=sublist_length)
            
            # remove the class from their list
            del student_preference_dict[student][0]

            # assign student to that session
            class_index = available_sessions_to_take.index(session_with_least_students)
            class_dict[student_choice][class_index].append(student)
            
            # put the class in the student's timetable
            dict_student_classschedules[student].update({student_choice : class_index})
            # go to the next student

   # Remove available class
    Remove_Full_Class(max_class_size,class_dict,available_classes)

    # Assign classes for students with no preference
    # calculate the smallest session instead of smallest sum of class?
    #'''
    for i in range(num_workshop_rounds):
        for student in students_with_no_preference:
            # get the list of class taken by current student, remove them from the available choice 
            classes_already_taken = dict_student_classschedules[student]
            class_to_choose_from = [x for x in available_classes if x not in classes_already_taken]

            if not class_to_choose_from or not available_classes:
                print("No available classes found! Find another solution for this")

            # Create a temporary copy of the dict, remove the class that's already taken by student.
            new_class_dict = {key: value for key, value in class_dict.items() if key in class_to_choose_from}
            
            # Calculate the sum of total students for each workshop in the new dict
            total_student_in_each_workshop = {key: sum(len(sublist) for sublist in value) for key, value in new_class_dict.items()}

            # Pick the class with the least number of students in total
            class_with_least_students = min(total_student_in_each_workshop, key=total_student_in_each_workshop.get)
            chosen_workshop_sessions = class_dict[class_with_least_students]

            # access the index of classes students are taking
            classindexes = list(dict_student_classschedules[student].values())
            # remove those from consideration in this min()???
            classindexes.sort()
            available_sessions_to_take = chosen_workshop_sessions.copy()

            for index in classindexes:
                available_sessions_to_take[index] = None

            # take the session with least number of people
            session_with_least_students = min(filter(lambda x: x is not None, available_sessions_to_take), key=sublist_length)
            
            while len(session_with_least_students) >= max_class_size:
                total_student_in_each_workshop.pop(class_with_least_students) # dict of class name : number of total students
                
                if not total_student_in_each_workshop:
                    print("WARNING! No more available class for", student,"\n",class_with_least_students,"will exceed the maximum")
                    # DO SOMETHING
                    break
                # Pick the class with the least number of students in total
                class_with_least_students = min(total_student_in_each_workshop, key=total_student_in_each_workshop.get)
                chosen_workshop_sessions = class_dict[class_with_least_students]

                available_sessions_to_take = chosen_workshop_sessions.copy()

                for index in classindexes:  #index is already previously calculated
                    available_sessions_to_take[index] = None
                
                session_with_least_students = min(filter(lambda x: x is not None, available_sessions_to_take), key=sublist_length)
            
            # Assign student to the session
            class_index = available_sessions_to_take.index(session_with_least_students)
            class_dict[class_with_least_students][class_index].append(student)
            
            # put the class in the student's timetable
            dict_student_classschedules[student].update({class_with_least_students : class_index})
        
        # loop through the available class and get rid of the actual full one here
        Remove_Full_Class(max_class_size,class_dict,available_classes)
    #'''
    
    # Print the summary
    print("**********CLASS SUMMARY*****************")
    for key, value in class_dict.items():
        print(key)
        for sublist in value:
            print(len(sublist),":",sublist)
    
    '''
    print("\n\n**********STUDENT SCHEDULE*****************")
    for key, value in dict_student_classschedules.items():
        print(key)
        print(len(value),":",value)
    '''
    return dict_student_classschedules  
    
# Rearrange student's schedule based on the order of the class
def Rearrange_Student_Schedule(student_schedule_dict):
    # student_preference_dict: student_A:{[]:2,[],10}
    converted_dict = {}
    for key, value in student_schedule_dict.items():
        sorted_subdict = sorted(value.items(), key=lambda x: x[1])
        converted_list = [item[0] for item in sorted_subdict]
        converted_dict[key] = converted_list

    #for k,v in converted_dict.items():
        #print(k,":",v)
    return converted_dict
    
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
Cells_Cleanup(2, 15, None, 19)

student_schedule_dict = Sort_Student_to_Workshop_by_Preference()
student_schedule = Rearrange_Student_Schedule(student_schedule_dict)

Excel_Update_Student_Schedule(student_schedule)