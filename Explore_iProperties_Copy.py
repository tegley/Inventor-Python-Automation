#Import Statements 
import win32com.client as wc
import time

def main(): #Control Inventor and execution of functions
	global inventor
	try: #Attempt to connect Inventor to Python
		inventor = wc.GetActiveObject('Inventor.Application')
		print(f"Inventor is connected to Python!")
	except:
		print(f"Unable to connect Python to Inventor :(")
		exit()
	
	#Collect user choice and select active document
	choice = int(input("Would you like to access property sets (1), update iProperties for a several open parts (2), update iProperties for an assembly (3), or determine all the iProperties of one part (4) -> "))
	inventor_part = inventor.ActiveDocument

	#Execute appropriate function based on user choice
	if choice == 1:
		Access_iProperties(inventor_part)
	elif choice == 2:
		Several_Parts_iProperties()
	elif choice == 3:
		Assembly_iProperties()
	elif choice == 4:
		Print_All_iProperties(inventor_part)

def Access_iProperties(inventor_part): #Access Property Sets for iProperties; 1 part must be open
	#Access property sets
	inventor_design_tracking_properties = inventor_part.PropertySets.Item("Design Tracking Properties")
	inventor_summary_information = inventor_part.PropertySets.Item("Inventor Summary Information")
	inventor_document_summmary_information = inventor_part.PropertySets.Item("Inventor Document Summary Information")
	inventor_user_defined_properties = inventor_part.PropertySets.Item("Inventor User Defined Properties")	

	#Display property sets
	print(f"Property Set 1: {inventor_design_tracking_properties.DisplayName}")
	print(f"Property Set 2: {inventor_summary_information.DisplayName}")
	print(f"Property Set 3: {inventor_document_summmary_information.DisplayName}")
	print(f"Property Set 4: {inventor_user_defined_properties.DisplayName}")

	#Update part number	
	part_number = inventor_design_tracking_properties.Item("Part Number")
	part_number.Value = "PCL - TOGETHER WE BUILD SUCCESS"

	#Add today's date
	date_checked = inventor_design_tracking_properties.Item("Date Checked")
	todays_date = Todays_Date()
	date_checked.Value = todays_date

def Several_Parts_iProperties(): #Access open part files and update iProperties; 1 or more parts must be open (no assembly files)
	inventor_documents = inventor.Documents #Determine all documents that are currently open
	total_docs_open = int(inventor_documents.Count)
	print(f"Total number of part documents open: {total_docs_open}")
	total_docs_open += 1
	for i in range(1, total_docs_open): #Iterate through the open documents
		current_document = inventor_documents.Item(i)
		date_checked = current_document.PropertySets.Item("Design Tracking Properties").Item("Date Checked") #Change the date checked to today's date
		todays_date = Todays_Date()
		date_checked.Value = todays_date
	print(f"All dates were updated to reflect today's date: {date_checked.Value}")

def Assembly_iProperties(): #Access parts from an assembly file and update their individual iProperties; 1 assembly must be open
	assembly = inventor.ActiveDocument #Configure assembly file
	
	#Access and print assembly document name
	assembly_name = assembly.DisplayName
	print(f"Assembly file is: {assembly_name}")
	
	parts_in_assembly = assembly.AllReferencedDocuments #Locate all referenced documents (parts and sub-assemblies used in the assembly)
	number_of_parts = parts_in_assembly.Count #Total number of part files in the assembly
	print(f"Total number of parts in assembly: {number_of_parts}")
	number_of_parts += 1

	for x in range(1, number_of_parts): #Iterate through all documents referenced in the assembly
		current_part = parts_in_assembly.Item(x) #Access current file
		current_part_name = str(current_part.DisplayName)
		
		file_type = current_part_name[-4:] #Use negative indexing to pull the file type (determine whether current file is a part or an assembly)
		
		if file_type == ".ipt": #Perform iProperty updates for part files
			#Update today's date
			print(current_part.DisplayName)
			date_checked = current_part.PropertySets.Item("Design Tracking Properties").Item("Date Checked") #Change the date checked to today's date
			todays_date = Todays_Date()
			date_checked.Value = todays_date
			print("Date updated!")

		elif file_type == ".iam": #Skip over sub-assembly files
			print(f"Sub Assembly: {current_part}")

		else:
			continue

def Print_All_iProperties(inventor_part): #Print all the iProperties of one part; 1 part must be open
	#The commented code below can be used to determine options for each property set
	
	# #Access property sets
	# inventor_design_tracking_properties = inventor_part.PropertySets.Item("Design Tracking Properties")
	# inventor_summary_information = inventor_part.PropertySets.Item("Inventor Summary Information")
	# inventor_document_summmary_information = inventor_part.PropertySets.Item("Inventor Document Summary Information")
	# inventor_user_defined_properties = inventor_part.PropertySets.Item("Inventor User Defined Properties")	

	# #Print options for each property sets
	# print("PROPERTY SETS:")
	# print(dir(inventor_design_tracking_properties))
	# print(dir(inventor_summary_information))
	# print(dir(inventor_document_summmary_information))
	# print(dir(inventor_user_defined_properties))

	
	inventor_property_sets = inventor_part.PropertySets #Overarching property sets
	property_sets_count = inventor_property_sets.Count #Total number of property sets
	property_sets_count += 1
	
	#Print all property sets, the number of text boxes of each, and the value of each text box
	for x in range(1, property_sets_count):
		display_name = inventor_property_sets.Item(x).Name
		count = int(inventor_property_sets.Item(x).Count) #Total number of text boxes within the property set
		print(f'{display_name}: {count} text boxes') #Print property set and number of text boxes
		count += 1
		for y in range(1, count):
			print(inventor_property_sets.Item(x).Item(y).Name) #Print corresponding text box names
			try: #Attempt to print their value if it isn't blank
				print(inventor_property_sets.Item(x).Item(y).Value)
			except:
				print("No value")
		print("\n")

	try:
		#Determine and display sheet metal properties
		length = float(inventor_part.ComponentDefinition.FlatPattern.Length) / 2.54 #Convert length of plate to inches
		width = float(inventor_part.ComponentDefinition.FlatPattern.Width) / 2.54 #Convert width of plate to inches
		thickness = float(inventor_part.ComponentDefinition.Thickness.Value) / 2.54 #Convert plate thickness to inches
		print("SHEET METAL PROPERTIES")
		print(f'Length: {length}')
		print(f'Width: {width}')
		print(f'Thickness: {thickness}')

	except:
		print("This is not a sheet metal part!")

#Secondary function
def Todays_Date(): #Determine todays date using the time module
	current_time = time.time() #Current time in seconds
	time_object = time.localtime(current_time) #Current time (converted)
	
	#Format date
	year = int(time_object.tm_year)
	month = int(time_object.tm_mon)
	day = int(time_object.tm_mday)
	time_string = f"{month}/{day}/{year}" 
	
	return(time_string) #Send todays date back to calling function

main() #Run the program