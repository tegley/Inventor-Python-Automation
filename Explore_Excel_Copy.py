#Import Statements
import openpyxl
import warnings

def main():
    #Ask for file path and text to be typed onto cell A1
    user_inputted_path = input("Enter file path of XLSM file: ")
    correct_file_path = Get_File_Path(user_inputted_path)
    users_text = input("What would you like to type on cell A1 -> ")
    Type_User_Text(correct_file_path, users_text)

#Primary Functions
def Get_File_Path(user_inputted_path): #Convert the user-inputted file path to a string that Python can use to interpret the file location
	#Necessary Unicode characters
	backslash = '\u005C' #Unicode for \
	double_quote = '\u0022' #Unicode for "

	#Identify folders based
	folders = user_inputted_path.split(backslash)
	end = len(folders) - 1    

	#Remove double quotes
	folders[0] = folders[0].replace(double_quote,'')
	folders[end] = folders[end].replace(double_quote,'')

	#Determine where the for loop should end
	total = len(folders) - 1

	#Add in backslashes
	for i in range(0, total):
		placeholder = folders[i]
		folders[i] = placeholder + backslash

	#Convert finalized file path to a string    
	user_inputted_path = ''.join(folders)
	return(user_inputted_path) 

def Type_User_Text(path, users_input):
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl') #Ignore openpyxl warning

    #Open XLSM file
    excel = openpyxl.load_workbook(path, read_only=False, keep_vba=True)
    requisition = excel['REQUISITION']
    
    #Access cell A1 and input users desired text
    requisition['A1'].value = str(users_input)
    cell = Increment_Row('A1','A', 1)
    requisition[cell].value = "Python is cool"
    
    #Save Excel file
    excel.save(path)

#Secondary function
def Increment_Row(cell, column, amount):
    letter_and_number = list(cell.partition(column)) #
    letter_and_number.pop(0)
    
    column = letter_and_number[0] 
    letter_and_number[1] = str(int(letter_and_number[1]) + amount)
    
    cell = ''.join(letter_and_number)
    return(cell)

main() #Run the program