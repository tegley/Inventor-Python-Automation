#Length/Width/Thickness
# inventor_part_w = inventor.ActiveDocument.ComponentDefinition.FlatPattern.Width #Width of sheet metal part 
# inventor_part_l = inventor.ActiveDocument.ComponentDefinition.FlatPattern.Length #Length of sheet metal part
# inventor_part_t = inventor.ActiveDocument.ComponentDefinition.Thickness.Value #Thickness of sheet metal part

#Helpful API Stuff
#Parameter Object, Value Property
#kDefaultDisplayLengthUnits
#DimensionConstraints.AddTwoPointDistance Method
#ExtrudeFeaturesObject
#ExtrudeDefinitionObject
#Sketch from Face Silhouette API Sample/SketchCircles Object
#print(dir(inventor_part))
#PlanarSketch

#Unicode
#For 00 characters, use \u00
#Ex. \u0022 for "
#For 1F characters, use \U000
#Ex. \U0001F600 for 😀

#Useful code
#Break, Continue, and Pass for loops
#Try and except

#Debug names
# inventor_summary_info = inventor_part.PropertySets
# print(dir(inventor_summary_info)
# count = inventor_summary_info.Count + 1
# for x in range(1,count):
# 	print(inventor_summary_info.Item(x).Name)

#Advanced debug
# def Run_iProperties(inventor_part):
# 	inventor_summary_info = inventor_part.PropertySets
# 	property_sets_count = inventor_summary_info.Count + 1
# 	for x in range(1, property_sets_count):
# 		name = inventor_summary_info.Item(x).Name
# 		count = int(inventor_summary_info.Item(x).Count)
# 		print(f'{name}: {count} text boxes')
# 		count += 1
# 		for y in range(1, count):
# 			print(inventor_summary_info.Item(x).Item(y).Name)
# 		print("\n")
	
# 	inventor_design_info = inventor_part.PropertySets.Item("Design Tracking Properties")
# 	part_number = inventor_design_info.Item("Part Number")
# 	part_number.Value = "PCL is awesome!"
# 	print("Check part number!")

#print(current_document.DisplayName) #Displays the part name
#print(current_document.FullFileName) #Displays the full file path

#Scrapped Code
#    print(dim_number[0])
#    print(type(dim_number))
#    print(dim_number)
#    print(notation)


#Access rows code working
	# rows = 11
	# for z in range(1, rows): #Print out cell values
	# 	current_row = parts_list.PartsListRows.Item(z)
	# 	for y in range(1,columns):
	# 		current_column = parts_list.PartsListColumns.Item(y).Title
	# 		print(current_column)
	# 		cell = current_row.Item(y)
	# 		print(cell)
	# 		if cell.Value == "Construction_Worker_With_Hat_Final":
	# 			cell.Value = "I figured it out!"

#Iterate through BOM (rows first) (Line by line)
	# for a in range(1, rows): #Print out/access cell values
	# 	current_row = bom.PartsListRows.Item(a)
	# 	for b in range(1,columns):
	# 		current_column = bom.PartsListColumns.Item(b).Title
	# 		print(current_column)
	# 		cell = current_row.Item(b)
	# 		print(cell)


#Iterate through BOM (columns first) (Fill out each column)
	# for a in range(1, columns):
	# 	current_column = bom.PartsListColumns.Item(a)
	# 	title = current_column.Title
	# 	print(title)
	# 	for b in range(1,rows):
	# 		current_row = bom.PartsListRows.Item(b).Item(a).Value
	# 		print((current_row))

#VBA Code for BOM
# 'Declare local variables
# Dim oDrawDoc As DrawingDocument
# Dim oSheet As Sheet
# Dim oDrawingView As DrawingView
# Dim oPlacementPoint As Point2d
# Dim oPartsList As PartsList 
# Dim oBorder As Border
# Dim oStyle As String
# Dim oPlaceX As Double
# Dim oPlaceY As Double
# Dim oPartsListDeltaX As Double

# 'initialize variables
# oDrawDoc = ThisApplication.ActiveDocument
# oSheet = oDrawDoc.ActiveSheet
# oDrawingView = oSheet.DrawingViews(1)
# oBorder = oSheet.Border
# oStyle = "Custom Parts List (ANSI)"

# 'user input - select parts list style
# 'oStyle = InputListBox("Choose Parts List Style", MultiValue.List("PartsListStyle"), "Custom Parts List (ANSI)", "Parts List Style", "List Prompt") 

# 'create parts list
# oPartsList = oSheet.PartsLists.Add(oDrawingView, oBorder.RangeBox.MinPoint)
# oPartsList.Style = oDrawDoc.StylesManager.PartsListStyles.Item(oStyle​)

# 'reposition parts list to top left
# oPartsListDeltaX = oPartsList.RangeBox.MaxPoint.X-oPartsList.RangeBox.MinPoint.X
# oPlaceX = oBorder.RangeBox.MinPoint.X + oPartsListDeltaX
# oPlaceY = oBorder.RangeBox.MaxPoint.Y
# oPlacementPoint = ThisApplication.TransientGeometry.CreatePoint2d(oPlaceX,oPlaceY)
# oPartslist.position = oPlacementPoint
# End Sub

#Code to format inches
# def main():
# 	s = Format_Final_Display("0.75","\u0022")
# 	print(s)

# def Format_Final_Display(item, inch): #Format the inputted string based on the 
# 	mixed_number = item.split('.') #Differentiate between the whole number and the fractional part
# 	print(mixed_number)
	
# 	if int(mixed_number[0]) != 0: #Determine whether the inputted item has a whole number component
# 		whole_number = str(mixed_number[0])
# 	else:
# 		whole_number = 0
	
# 	if int(mixed_number[1]) != 0: #Determine whether the inputted item has a fractional component
# 		fraction = (float('.' + mixed_number[1])).as_integer_ratio()								
# 		numerator = str(fraction[0])
# 		denominator = str(fraction[1])	
# 	else:
# 		fraction = 0
	
	
# 	if whole_number != 0 and fraction != 0.0:  #Case for mixed number (ex. 1.25)
# 		finalized_item = f"{whole_number} {numerator}/{denominator}{inch}"
		
# 	elif whole_number == 0 and fraction != 0.0: #Case for only fractional part (ex. 0.25)
# 		finalized_item = f"{numerator}/{denominator}{inch}"
		
# 	elif whole_number != 0 and fraction == 0.0: #Case for only whole number part (ex. 1)
# 		finalized_item = f"{whole_number}{inch}"
	
# 	else:
# 		finalized_item = f"0{inch}" #Continue run in case of failed instances

# 	return(finalized_item)

# main()

#Edit/Save XLSM files
# As it is 2021, get_sheet_by_name is deprecated and raises an DeprecationWarning with the following message: Call to deprecated function get_sheet_by_name 
# (Use wb[sheetname]).

# The following snippet can be used in order to not raise the warning.

# from openpyxl import load_workbook

# file_path = 'test.xlsx'

# wb = load_workbook(file_path)

# ws = wb['SHEET_NAME']  # or wb.active

# ws['G6'] = 123

# wb.save(file_path)

#parts[0] = chr(ord(parts[0]) + 1) #Iterate through Excel columns
#


#Useful links
#https://pypi.org/
#https://stackoverflow.com/questions/69547919/f2py-exe-is-somewhere-but-the-directory-isnt-on-path
#https://www.unicode.org/Public/UCD/latest/ucd/NamesList.txt
#https://realpython.com/python-string-contains-substring/
#https://www.geeksforgeeks.org/python-reading-excel-file-using-openpyxl-module/
# thing = '\u005C'
# print(thing)
# print("/")
#https://stackoverflow.com/questions/64420348/ignore-userwarning-from-openpyxl-using-pandas

# def Update_String(cell, column):
# 	letter_and_number = list(cell.partition(column)) #
# 	letter_and_number.pop(0)
	
# 	letter_and_number[0] = chr(ord(letter_and_number[0]) + 1) #Update column
# 	column = letter_and_number[0] #Store column as a string for use in next iteration
# 	letter_and_number[1] = str(int(letter_and_number[1]) + 1) #Update row
	
# 	cell = ''.join(letter_and_number) #Generate correct string from list 
# 	print(cell, column)
    
# Update_String("A26", "A")
# Output: B27 B


# def Increment_Column(cell, column): #Increment the row of a cell in Excel
#     letter_and_number = list(cell.partition(column))
#     letter_and_number.pop(0)
    
#     letter_and_number[0] = chr(ord(letter_and_number[0]) + 1) #Increment the column; row = letter_and_number[1]
    
#     cell = ''.join(letter_and_number)
#     print(cell)


#Fill out req via columns (up and down)
	# bom_column = 3 #Identify where the Description column is located on the BOM
	# req_cell = 'D25' #The cell where the Description line starts in the Requisition
	# for num_1 in range(1, total_bom_rows): #Fill out the rows of the Description column
	# 	bom_row = bom.PartsListRows.Item(num_1).Item(bom_column).Value
	# 	req_cell = Increment_Row(req_cell,'D')
	# 	requisition[req_cell].value = (bom_row)
	
	# bom_column = 4 #Identify where Material column is located on the BOM
	# req_cell = 'E25' #The cell where the Mat Spec column starts in the Requisition
	# for num_2 in range(1, total_bom_rows): #Fill out the rows of the mat spec column
	# 	bom_row = bom.PartsListRows.Item(num_2).Item(bom_column).Value
	# 	req_cell = Increment_Row(req_cell,'E')
	# 	requisition[req_cell].value = (bom_row)


# #Import Statements 
# import win32com.client as wc #Using this library to control Inventor via Python
# # import os #Using this library to start Inventor via the operating system
# # import time #Using this library to control when actions occur and determine the current date
# # import openpyxl
# # import warnings

# inventor = wc.GetActiveObject('Inventor.Application') #Connect Inventor to Python
# this = inventor.Documents.VisibleDocuments.Item(2).DocumentType
# print((this))

# sheet = this.ActiveSheet

# view = sheet.DrawingViews(1)

# #Identify border and set placement points
# border = sheet.Border
# placement_1 = border.RangeBox.MaxPoint

# #Identify parts list set-ups that are going to be used
# bom_style = "Custom Parts List (ANSI)"

# bom = sheet.PartsLists.Add(view, placement_1, 46595)

# print((sheet))


#Notes for 7/23/24
#Can - find way to convert to flat pattern via code
#Remeber to convert to inches!
#Pipe is feet (UOM) #Pipe naming is by nominal pipe size (1) not length (10.5) #Round by the foot (eg 50.1 is 51)
#Pipe convert from inches to feet, round to the nearest 1/8", display everything in decimals
#Consolidation - do it by the item desciption and see how it goes (string comparison testing)
#Account for varied material specificiations

# #Fix this - it rounds down!
# feet  = float(13 / 12)
# rounded = int(feet / 0.125) * 0.125
# if rounded < feet:
#     rounded += 0.125
# print(rounded)


# import openpyxl
# backslash = '\u005C' #Unicode for \
# path = f"N:{backslash}District Office{backslash}Engineering{backslash}Technical Documents{backslash}Tim Egley{backslash}Inventor{backslash}Inventor Python Automation{backslash}629 Copy For Programming Use.xlsm"

# excel = openpyxl.load_workbook(path, read_only=False, keep_vba=True)
# print("Names in current sheet:", excel.sheetnames)
# requisition = excel['REQUISITION']

# prior_material_cell = "D26"
# prior_material = requisition[prior_material_cell].value
# print(prior_material)
# for scan in range(1, total_bom_rows):
# 		qty_counter_dummy += 1
# 		a = scan - 1
# 		if prior_bom_row == scan:
# 			continue		
# 		if consolidation_check[a] == True:
# 			continue	
# 		else:
# 			next_description = bom.PartsListRows.Item(scan).Item(description_column).Value
# 			next_material = bom.PartsListRows.Item(scan).Item(material_column).Value	
# 			if prior_description == next_description and prior_material == next_material:

				
				
				
				
# 				print("Prior quantity counter:", prior_qty_counter)
# 				print("Quantity dummy:", qty_counter_dummy)
			
# 				add_this = qty[qty_counter_dummy]
# 				consolidation_check[qty_counter_dummy] = True
		
# 				if uom == 'LF':
# 					initial = qty[prior_qty_counter]
# 					num += add_this + initial
# 					print("Initial", initial)
# 					print("Add this", add_this)

# 				elif uom == 'EA':
# 					initial = bom.PartsListRows.Item(scan).Item(2).Value
# 					requisition[quantity_cell_excel].value = add_this + initial

#Below this line is code that will be deleted!
# def Final_Scan(bom, prior_description, prior_material, total_bom_rows):
# 	local_counter = 0
# 	for scan in range(1, total_bom_rows):
# 		next_description = bom.PartsListRows.Item(scan).Item(3).Value
# 		next_material = bom.PartsListRows.Item(scan).Item(4).Value		
# 		if prior_description == next_description and prior_material == next_material:
# 			consolidation_check[local_counter] = True

# def Consolidate_Items(bom, requisition, total_bom_rows, quantity_cell_excel, bom_row, prior_bom_row, final, exists, prior_material, prior_description, quantity_cell_bom, prior_qty_counter):
# 	print("Quantity List:", qty_list)
# 	delete = []
# 	while exists == True:
# 		exists, counter = Scan(bom, bom_row, prior_bom_row, prior_description, prior_material, total_bom_rows)
# 		if exists == True:
# 			delete.append(counter)
# 			value = qty_list[counter] * float(quantity_cell_bom)
# 			print("One value:", value)
# 			final += value
# 			# consolidation_check[counter] = True
# 			# print("Check list:",consolidation_check)
# 		elif exists == False:
# 			break
# 	print("Final value post consolidation:", final)
# 	requisition[quantity_cell_excel].value = final		
	

# 	#Edit and fix list
# 	delete.reverse()
# 	delete_amount = len(delete)
# 	qty_reset = len(qty_list) - delete_amount + 1

# 	qty_list[prior_qty_counter] = final
# 	print("Pre deletion:",qty_list)	
# 	for d in range(0, delete_amount):
# 		qty_list[delete[d]] = 999.123
	
# 	print("999.123 Switch:", qty_list)

# 	for e in range(0, qty_reset):
# 		if qty_list[e] == 999.123:
# 			del qty_list[e]
# 		else:
# 			continue
# 	print("Final Quantiy (Post Deletion):", qty_list)

# 	return

# if uom == 'LF':
			# 	initial = qty_list[prior_qty_counter]
			# elif uom == 'EA':
			# 	initial = bom.PartsListRows.Item(prior_bom_row).Item(2).Value
		
			# final = 0
			# final += initial
			# print("Original item value:",final)

			# Consolidate_Items(bom, requisition, total_bom_rows, quantity_cell_excel, bom_row, prior_bom_row, final, exists, prior_material, prior_description, quantity_cell_bom, prior_qty_counter)
			# Final_Scan(bom, prior_description, prior_material, total_bom_rows)

# def Type_Requisition_Item(bom, requisition, bom_row, qty_counter, req_cell):
# 		print("List as it appears in the Type Requisiton Item Function:", qty_list)
# 		bom_column = 2 #Reset BOM column before next loop iteration
# 		if qty_list[qty_counter] != 0.0: #Account for situations (pipe) where items are purchased by the linear foot
# 			quantity_cell_bom = float(bom.PartsListRows.Item(bom_row).Item(bom_column).Value) #Quantity
# 			total_linear_feet = quantity_cell_bom * float(qty_list[qty_counter]) #Calculate total linear feet of pipe
# 			requisition[req_cell].value = str(total_linear_feet)
# 			quantity_cell_excel = str(req_cell)

# 			req_cell, column, row = Increment_Column(req_cell,'B', 1) #Increment the column from 'B' to 'C'

# 			uom = 'LF'
# 			requisition[req_cell].value = uom #Set UOM as 'LF' (Linear Feet)
		
# 		else: #Account for majority of situations where items are purchased by the total amount needed
# 			quantity_cell_bom = bom.PartsListRows.Item(bom_row).Item(bom_column).Value #Quantity
# 			requisition[req_cell].value = (quantity_cell_bom)
# 			quantity_cell_excel = req_cell

# 			req_cell, column, row = Increment_Column(req_cell,'B', 1) #Increment the column from 'B' to 'C'

# 			uom = 'EA'
# 			requisition[req_cell].value = uom #Set UOM as 'EA' (Each)
		
# 		req_cell, column, row = Increment_Column(req_cell, column, 1) #Increment the column from 'C' to 'D'
# 		bom_column += 1 #Locate Description in BOM
		
# 		description_cell_bom = bom.PartsListRows.Item(bom_row).Item(bom_column).Value #Item Description cell in Requisition
# 		requisition[req_cell].value = (description_cell_bom) 

# 		req_cell, column, row = Increment_Column(req_cell, column, 1) #Increment the column from 'D' to 'E'
# 		bom_column += 1 #Locate Material in BOM

# 		material_cell_bom = bom.PartsListRows.Item(bom_row).Item(bom_column).Value #Mat Spec cell in Requisition
# 		requisition[req_cell].value = (material_cell_bom)

# 		req_cell = str('B' + str(int(row) + 1)) #Increment the row and reset the column to the QTY column before next for loop iteration
# 		prior_qty_req_cell = str('B' + str(int(row) + 1))

# 		return(req_cell, prior_qty_req_cell)


# CAP BOLT - 3/4", 2-1/2" LG