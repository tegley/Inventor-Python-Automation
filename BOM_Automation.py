#Import Statements 
import win32com.client as wc #Using this library to control Inventor via Python (pywin32)
import os #Using this library to start Inventor via the operating system
import time #Using this library to control when actions occur and determine the current date
import openpyxl #Using this library to fill out the Material Requisition Excel file
import warnings #Using this library to disable to warning message from openpyxl

#Primary Functions
def main(): #Control Inventor and execution of functions
	global inventor #Define inventor (application object) and various lists as a global variables that can be used anywhere in the code	
	inventor = wc.GetActiveObject('Inventor.Application') #Connect Inventor to Python

	#Differentiate between assembly document and drawing document
	document_1 = int(inventor.Documents.VisibleDocuments.Item(1).DocumentType)
	
	if document_1 == 12291: #kAssemblyDocumentObject
		assembly_document = inventor.Documents.VisibleDocuments.Item(1)
		drawing_document = inventor.Documents.VisibleDocuments.Item(2)
	elif document_1 == 12292: #kDrawingDocumentObject
		drawing_document = inventor.Documents.VisibleDocuments.Item(1)
		assembly_document = inventor.Documents.VisibleDocuments.Item(2)
	else:
		print("Correct documents were not open and/or incorrect documents were open. Program has crashed.")
		exit()

	print(f"Inventor is connected to Python!")
	print(f"Instructions:")
	print(f"Assembly and corresponding drawing with at least one view must be open")
	print(f"Material Requisition Excel file must be closed")
	
	path_check = input("Enter file path of Material Requisition to start automation: ")
	path, exists = Get_File_Path(path_check)	

	if exists == True: #If the file exists, run the program
		Edit_iProperties(assembly_document) #Execute Step 1
		bom = Create_BOM(drawing_document) #Execute Step 2
		Update_Material_Requisition(bom, path) #Execute Step 3

	elif exists == False:
		smile = "\U0001F600"
		print(f"Have a nice day {smile}") #End program

#Step 1
def Edit_iProperties(assembly_document): #Access parts from an assembly file and update their individual iProperties
	assembly_name = assembly_document.DisplayName 
	print(f"Assembly file is: {assembly_name}")
	parts_in_assembly = assembly_document.AllReferencedDocuments #Locate all parts and sub-assemblies used in the assembly)
	number_of_parts = parts_in_assembly.Count

	for x in range(number_of_parts, 0, -1): #Iterate through all documents referenced in the assembly from top to bottom
		current_part = parts_in_assembly.Item(x) #Access current file
		current_part_name = str(current_part.DisplayName)
		
		#Overarching iProperty sets
		inventor_design_tracking_properties = current_part.PropertySets.Item("Design Tracking Properties")
		inventor_summary_information = current_part.PropertySets.Item("Inventor Summary Information")
		inventor_document_summmary_information = current_part.PropertySets.Item("Inventor Document Summary Information")
		inventor_user_defined_properties = current_part.PropertySets.Item("Inventor User Defined Properties")
			
		#Pull description and status text boxes for BOM use
		description = inventor_design_tracking_properties.Item("Description")
		status = inventor_design_tracking_properties.Item("User Status")
		inch = "\u0022" #Unicode for "
				
		#For all files that are parts, edit iProperties and update the description according to PCL standards
		if ".ipt" in current_part_name:
			#Change date to todays date
			date_checked = inventor_design_tracking_properties.Item("Date Checked")
			todays_date = Todays_Date()
			date_checked.Value = todays_date

			#Pull part number amd store it as a list
			part_number = inventor_design_tracking_properties.Item("Part Number").Value
			part_number_string = str(part_number)
			part_words = list(part_number_string.split())
			
			#Make part number case insensitive
			part_words_casefold = list(part_number_string.split())
			list_length = len(part_words_casefold)
			for i in range(0, list_length):
				part_words_casefold[i] = str(part_words_casefold[i]).casefold()

			#Distingish between content library parts and Angel's sheetmetal parts
			document_subtype_name = inventor_design_tracking_properties.Item("Document SubType Name").Value

			if "Layout_1" in part_words_casefold: #Skip over the 'Layout_1' part
				continue
			
			elif document_subtype_name == "Modeling": #Account for parts in the Content Library
				
				#Update description of pipe in accordance to PCL standards; LF item
				if "pipe" in part_words_casefold: #Naming convention: ASME B36.10M Pipe (Diameter) - Schedule (schedule) - (length)

					#Pull 'Schedule' as a keyword
					pull = int(part_words.index("Schedule"))
					#Determine diameter
					pull -= 2
					diameter = str(part_words[pull])

					#Determine all components of the diameter
					based_on_diameter = list(part_number_string.partition(diameter)) #Split part number based on the diameter
					first_part = list(based_on_diameter[0].split()) #Contains everything in the part number before the diameter
					last_of_first = first_part.pop() #Is either the whole part of the mixed number or an irrelevant word

					try: #Determine whether the diameter is a mixed number or not
						num = int(last_of_first)
					except:
						num = str(last_of_first)
					
					if type(num) == int: #Format diameter accordingly
						final_diameter = f"{num} {diameter}{inch}"
					else:
						final_diameter = f"{diameter}{inch}"

					#Determine the pipe schedule based on the part number
					pull += 3
					schedule = str(part_words[pull])
					
					#Identify pipe as a 'linear feet' item
					length = float(part_words.pop()) #Length is the last item listed in the Part Number
					length /= 12 #Convert length from inches to feet
					status.value = str(length) + ";.ipt"
					
					#Create final description
					final_description = f"PIPE - {final_diameter}, SCH {schedule}, SMLS"					
					description.Value  = final_description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")
	
				#Update description of flanges in accordance to PCL standards; EA item
				elif "flange" in part_words_casefold: #Naming convention: 
					#Identify plate as an 'each' item
					status.value = str(0.0) + ";.ipt"

					#Pull 'Class' as a keyword
					pull = int(part_words.index("Class"))
					
					#Determine flange diamter
					pull += 2
					diameter = str(part_words[pull])
					
					#Determine flange class rating
					pull -= 1
					flange_class = str(part_words[pull])

					#Determine flange type
					flange_type = 'RFWN'

					#Create final description
					final_description = f"FLANGE - {diameter}, {flange_type}, {flange_class}"
					description.Value  = final_description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")

				#Update description of angle in accordance to PCL standards; LF item
				elif "L" in part_words: #Naming convention:
					#Pull "L" as a key word
					pull = int(part_words.index("L"))
					pull += 1
					dim_1 = str(part_words[pull])
					formatted_dim_1 = Format_Final_Display(dim_1)

					pull += 2
					dim_2 = str(part_words[pull])
					formatted_dim_2 = Format_Final_Display(dim_2)
					
					pull += 2
					thickness = str(part_words[pull])
					pull += 1
					item = str(part_words[pull]) #Determine if thickness is a mixed number or not
					
					if item == "-": #Would mean that thickiness is not a mixed number
						final_thickness = f"{thickness}{inch}"
					else: #Would mean that thickness is a mixed number
						final_thickness = f"{thickness} {item}{inch}"
						
					#Identify angles as a 'linear feet' item
					length = float(part_words.pop()) #Length is the last item listed in the Part Number
					length /= 12 #Convert length from inches to feet
					status.value = str(length) + ";.ipt"				
					
					final_description = f"ANGLE - {formatted_dim_1} x {formatted_dim_2} x {final_thickness}"
					description.Value  = final_description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")

				#Update description of cap bolts in accordance to PCL standards; EA item
				elif "UNC" in part_words:
					status.value = str(0.0) + ";.ipt"				
					
					unc = list(part_number_string.partition("UNC")) #Create a list based on the word "UNC"
					sizing = unc[0].split().pop() #Bolt diameter and threads per inch
					sizing_2 = unc[0].split().pop(-2) #Part of the sizing if it is a mixed number
					
					length = unc[2].split().pop(1) #Length is the last part of the final string item
					
					if length[-1] == ",": #Delete a comma (if it is there) through character iteration
						length = length[:-1]

					format_length = Format_Final_Display(length)
					
					sizing_check = sizing.split('-') #Remove dash and seperate threads/sizing
					sizing = sizing_check[0]
					threads = sizing_check[1]

					try: #For case that sizing is a mixed number
						sizing_check = int(sizing_2)
						final_description = f"CAP BOLT - {sizing_2} {sizing}{inch} - {threads} x {format_length} LG"

					except: #For case that sizing is a whole number or fraction
						sizing_check = str(sizing_2)
						final_description = f"CAP BOLT - {sizing}{inch} - {threads} x {format_length} LG"
					
					description.Value  = final_description #Create final description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")
 
				#Identify grating parts and enable the user to enter the final sizing accordingly
				elif "Smart_Grating" in part_words: 
					status.value = str(0.0) + ";.ipt"
					final_description = f'GRATING - Enter sizing'
					description.Value  = final_description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")

				else: #Account for other content library parts that have yet to be coded
					status.value = str(0.0) + ";.ipt"
					description.Value  = "Content Library Part Recognized: Code not yet developed"
					print(part_words)
					print("Content Library Part Recognized: Code not yet developed")

			elif document_subtype_name == "Sheet Metal": #Account for Sheet Metal parts that Angel has created
				#Update description for current part if it is a plate or angle tab
				if "plate" in part_words_casefold or "tab" in part_words_casefold or "toeboard" in part_words_casefold: #Only naming reqirement is for the word "plate" to be in the part number
					if current_part.ComponentDefinition.HasFlatPattern == False: #Create flat pattern (if not already previously created)
						create_flat_pattern_button = (inventor.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd"))
						create_flat_pattern_button.Execute()
						current_part.ComponentDefinition.Unfold()
					
					#Identify plate as an 'each' item
					status.value = str(0.0) + ";.ipt"

					#Determine length of plate
					length = float(current_part.ComponentDefinition.FlatPattern.Length) / 2.54 #Convert length of plate to inches
					rounded_length = str(Round_To_Nearest_Eighth(length)) #Round length to nearest 1/8"
					format_length = Format_Final_Display(rounded_length) #Format length

					#Determine width of plate
					width = float(current_part.ComponentDefinition.FlatPattern.Width) / 2.54 #Convert width of plate to inches
					rounded_width = str(Round_To_Nearest_Eighth(width)) #Round width to nearest 1/8"
					format_width = Format_Final_Display(rounded_width) #Format width

					#Determine length of plate
					thickness = float(current_part.ComponentDefinition.Thickness.Value) / 2.54 #Convert length of plate to inches
					rounded_thickness = str(Round_To_Nearest_Eighth(thickness)) #Round thickness to nearest 1/8"
					format_thickness = Format_Final_Display(rounded_thickness) #Format thickness
					
					#Calculate the area in square feet
					area = round(((length * width) / 144), 4) #Round area to 4 decimal places
					area = str(area)
					
					#Create final description and ensure that the length is always greater than the width
					if length >= width:
						final_description = f"PLATE - {format_thickness} THK x {format_length} LG x {format_width} W ({area} SQ FT)"
					elif length < width:
						final_description = f"PLATE - {format_thickness} THK x {format_width} LG x {format_length} W ({area} SQ FT)"
					
					description.Value = final_description
					print(f"Part Name: {current_part_name}; Final Description {final_description}")

				else:
					description.Value = "Unique part for Angel recognized: Code not yet developed"
					status.value = str(0.0) + ";.ipt"
					print(f"{current_part}: Unique part for Angel recognized: Code not yet developed")

		elif ".iam" in current_part_name or ".dwg" in current_part_name:
			description.Value = "Sub-assembly"
			status.value = str(0.0) + ";.iam"
			print(f"Part Name: {current_part_name}; Sub-assembly")
			continue

		else:
			print("Unique file or code not created")
			continue

#Step 2
def Create_BOM(drawing_document): #Create a BOM on the drawing file
	#Identify drawing sheet and view
	sheet = drawing_document.ActiveSheet
	view = sheet.DrawingViews(1)

	#Identify border and set placement points
	border = sheet.Border
	placement_1 = border.RangeBox.MaxPoint
	placement_2 = border.RangeBox.MinPoint

	#Identify parts list set-ups that are going to be used
	bom_style = "Custom Parts List (ANSI)"
	nozzle_schedule_style = "Custom Nozzle Schedule Parts List (ANSI)"

	bom = sheet.PartsLists.Add(view, placement_1, 46595)
	nozzle_schedule = sheet.PartsLists.Add(view, placement_2, 46595)
	
	bom.Style = drawing_document.StylesManager.PartsListStyles.Item(bom_style)
	nozzle_schedule.Style = drawing_document.StylesManager.PartsListStyles.Item(nozzle_schedule_style)

	
	#Calculate the total number of columns and set variables for column deletion
	columns = int(nozzle_schedule.PartsListColumns.Count) + 1 
	delete_1 = 1
	delete_2 = 1

	#Iterate through columns of the Nozzle Schedule and update accordingly
	for n in range(1, columns):
		current_column = nozzle_schedule.PartsListColumns.Item(n)
		column_name = current_column.Title #Store column name as a string to enable apporpirate changes via the following if-else structure
		if column_name == "EXT PROJ (A)":
			current_column.Title = "EXT PROJ"
		elif column_name == "INT PROJ (B)":
			current_column.Title = "INT PROJ"	
		elif column_name == "PIPE LENGTH":
			current_column.Title = "PIPE LG"
		elif column_name == "REPAD T1":
			current_column.Title = "REPAD t1"
		elif column_name == "REPAD OD":
			current_column.Title = "REPAD O.D."
		elif column_name == "REPAD ID":
			current_column.Title = "REPAD I.D."
		elif column_name == "FILLET WELD SIZE (C)":
			current_column.Title = "FILLET WELD SIZE"	
		elif column_name == "FILLET WELD SIZE (D)":
			delete_1 = n		
		elif column_name == "FILLET WELD SIZE (E)":
			delete_2 = n - 1
		else: #Keep columns that are already correct intact
			continue
	
	#Delete columns for Filet Weld Size (D) and Filet Weld Size (E)
	nozzle_schedule.PartsListColumns.Item(delete_1).Remove()
	nozzle_schedule.PartsListColumns.Item(delete_2).Remove()
	
	#Identify the location of the top left of the border
	nozzle_x = border.RangeBox.MinPoint.X + (nozzle_schedule.RangeBox.MaxPoint.X - nozzle_schedule.RangeBox.MinPoint.X)
	nozzle_y = border.RangeBox.MaxPoint.Y
	top_left = inventor.TransientGeometry.CreatePoint2d(nozzle_x, nozzle_y)

	#Reposition the nozzle schedule parts list to the top left of the sheet
	nozzle_schedule.Position = top_left
	
	return(bom)

#Step 3
def Update_Material_Requisition(bom, path):
	#Disable openpyxl warning
	warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
	
	#Open excel file in 'read' mode and identify Requisition sheet
	excel = openpyxl.load_workbook(path, read_only=False, keep_vba=True)
	print("Names in current sheet:", excel.sheetnames)
	requisition = excel['REQUISITION']

	total_bom_rows = bom.PartsListRows.Count + 1
	print(total_bom_rows)
	row = 25

	for counter in range(1, total_bom_rows):
		status_initial = str(bom.PartsListRows.Item(counter).Item(6).Value).split(";") #Status
		status_cell_bom = float(status_initial[0])
		file_type = status_initial[1]
		quantity_cell_bom = float(bom.PartsListRows.Item(counter).Item(2).Value) #Quantity

		if file_type == ".iam":
			continue
		
		elif file_type == ".ipt":
			bom_column = 2 #Reset BOM column before next loop iteration
			row += 1 #Increment row to the next line
			req_cell = str('B' + str(row)) #Set the cell to be the beginning of the next line of the material requisition
			
			if status_cell_bom != 0.0: #Account for situations (pipe) where items are purchased by the linear foot
				math = quantity_cell_bom * status_cell_bom
				requisition[req_cell].value = str(math)

				req_cell, column = Increment_Column(req_cell,'B') #Increment the column from 'B' to 'C'
				
				uom = 'LF'
				requisition[req_cell].value = uom #Set UOM as 'LF' (Linear Feet)
			
			else: #Account for majority of situations where items are purchased by the total amount needed
				quantity_cell_bom = int(quantity_cell_bom)
				requisition[req_cell].value = str(quantity_cell_bom)

				req_cell, column = Increment_Column(req_cell,'B') #Increment the column from 'B' to 'C'

				uom = 'EA'
				requisition[req_cell].value = uom #Set UOM as 'EA' (Each)
			
			req_cell, column = Increment_Column(req_cell, column) #Increment the column from 'C' to 'D'
			bom_column += 1 #Locate Description in BOM
			
			description_cell_bom = bom.PartsListRows.Item(counter).Item(bom_column).Value #Item Description cell in Requisition
			requisition[req_cell].value = (description_cell_bom) 

			req_cell, column = Increment_Column(req_cell, column) #Increment the column from 'D' to 'E'
			bom_column += 1 #Locate Material in BOM

			material_cell_bom = bom.PartsListRows.Item(counter).Item(bom_column).Value #Mat Spec cell in Requisition
			requisition[req_cell].value = (material_cell_bom)
	
	#Save changes to Excel file
	excel.save(path)

#Secondary Functions
def Todays_Date(): #Determine todays date using the time module (output = string)
	current_time = time.time() #Current time in seconds
	time_object = time.localtime(current_time) #Current time (converted)
	year = int(time_object.tm_year)
	month = int(time_object.tm_mon)
	day = int(time_object.tm_mday)
	time_string = f"{month}/{day}/{year}" #Format date
	return(time_string) #Send todays date back to calling function

def Round_To_Nearest_Eighth(num): #Round a number (input = float) to the nearest eighth (ouput = string)
	num = round(num, 4)
	rounded = int(num / 0.125) * 0.125
	if rounded < num:
		rounded += 0.125
	return(rounded)

def Format_Final_Display(item): #Format the inputted string based on the situation (ouput = string)
	mixed_number = item.split('.') #Differentiate between the whole number and the fractional part
	inch = "\u0022" #Unicode for "

	if int(mixed_number[0]) != 0: #Determine whether the inputted item has a whole number component
		whole_number = str(mixed_number[0])
	else:
		whole_number = 0
	
	if int(mixed_number[1]) != 0: #Determine whether the inputted item has a fractional component
		fraction = (float('.' + mixed_number[1])).as_integer_ratio()								
		numerator = str(fraction[0])
		denominator = str(fraction[1])	
	else:
		fraction = 0
	
	if whole_number != 0 and fraction != 0.0:  #Case for mixed number (ex. 1.25)
		finalized_item = f"{whole_number} {numerator}/{denominator}{inch}"
		
	elif whole_number == 0 and fraction != 0.0: #Case for only fractional part (ex. 0.25)
		finalized_item = f"{numerator}/{denominator}{inch}"
		
	elif whole_number != 0 and fraction == 0.0: #Case for only whole number part (ex. 1)
		finalized_item = f"{whole_number}{inch}"
	
	else:
		finalized_item = f"0{inch}" #Continue run in case of failed instances

	return(finalized_item)

def Increment_Column(cell, column): #Increment the column of the inputed Excel cell (ex. A1 to B1); (inputs and ouputs are all strings except for amount, which is an int)
	letter_and_number = list(cell.partition(column))
	letter_and_number.pop(0)
	letter_and_number[0] = chr(ord(letter_and_number[0]) + 1)
	
	column = letter_and_number[0]
	row = letter_and_number[1]
	
	cell = ''.join(letter_and_number)
	
	return(cell, column)

def Get_File_Path(path_check): #Convert the user-inputted file path to a string that Python can use to interpret the file location
	#Necessary Unicode characters
	backslash = '\u005C' #Unicode for \
	double_quote = '\u0022' #Unicode for "

	#Identify folders based
	folders = path_check.split(backslash)
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
	path_check = ''.join(folders)

	exists = os.path.isfile(path_check) #Check if the user entered file path exists
	return(path_check, exists)

main() #Run the program
