#Import Statements 
import win32com.client as wc

def main(): #Control Inventor and execution of functions
    global inventor

    #Attempt to connect Inventor to Python
    try:
        inventor = wc.GetActiveObject('Inventor.Application') #Connect Inventor to Python
        print(f"Inventor is connected to Python!")
    except:
        print(f"Unable to connect Python to Inventor :(")
        exit()

    #Ask for user which function they want to run
    choice = int(input("Would you like to count the amount of documents (1), create a new file (2 - Part; 3 - Assembly; 4 - Drawing), or draw a rectangle (5) -> "))

    #Execute the appropriate function
    if choice == 1:
        Count_Documents()
    elif choice == 2:
        part = New_Part()
        print(part)
    elif choice == 3:
        assembly = New_Assembly()
        print(assembly)
    elif choice == 4:
        drawing = New_Drawing()
        print(drawing)
    elif choice == 5:
        Rectangle_Creation()

def Count_Documents(): #Count amount of documents open and print their name; 1 file must be open
    documents = inventor.Documents
    count = int(documents.Count) #Document counter (including assembly documents)
    print(f"There are {count} documents open")
    count += 1
    for i in range(1,count):
        current_document = documents.Item(i) #How each document is stored in the computer
        name = current_document.DisplayName #Document name
        type = int(current_document.DocumentType) #Determine document type
        if type == 12290: #kPartDocumentObject
            final = "A part"
        elif type == 12291: #kAssemblyDocumentObject
            final = "An assembly"
        elif type == 12292: #kDrawingDocumentObject
            final = "A drawing"
        
        print(f"{name}: {final}") #Print document name and type

def New_Part(): #Create a new part file; Inventor must be open
	inventor_part = inventor.Documents.Add(12290, inventor.FileManager.GetTemplateFile(12290,8963)) #kPartDocumentObject, kEnglishSystemOfMeasure
	return(inventor_part)

def New_Assembly(): #Create a new assembly file; Inventor must be open
    inventor_assembly = inventor.Documents.Add(12291, inventor.FileManager.GetTemplateFile(12291,8963)) #kAssemblyDocumentObject, kEnglishSystemOfMeasure
    return(inventor_assembly)

def New_Drawing(): #Create a new drawing file; Inventor must be open
    inventor_drawing = inventor.Documents.Add(12292, inventor.FileManager.GetTemplateFile(12292,8963)) #kDrawingDocumentObject, kEnglishSystemOfMeasure
    return(inventor_drawing)
	
def Rectangle_Creation(): #Generate a rectangle based on user input; Inventor must be open
    #User input
    length_choice = str(input("Select length of rectangle -> "))
    length_choice += " in"
    height_choice = str(input("Select height of rectangle -> "))
    height_choice += " in" 
    extrusion_amount = str(input("Select extrusion amount -> "))
    extrusion_amount += " in" 

    inventor_part = New_Part() #Generate part file

    #Define the part file and corresponding parameters
    part_definition = inventor_part.ComponentDefinition
    user_parameters = part_definition.Parameters.UserParameters
    model_parameters = part_definition.Parameters.ModelParameters
    reference_parameters = part_definition.Parameters.ReferenceParameters
    geometry = inventor.TransientGeometry

    #Create a new sketch
    sketch_1 = part_definition.Sketches.Add(part_definition.WorkPlanes.Item(3)) #Puts sketch on the XY Plane

    #Add length and width as user parameters
    user_parameters.AddByExpression("Length_1", length_choice, 11266) #kDefaultDisplayLengthUnits
    user_parameters.AddByExpression("Height_1", height_choice, 11266) #kDefaultDisplayLengthUnits
    
    #Generate two-point centered rectangle
    shape_1 = sketch_1.SketchLines.AddAsTwoPointCenteredRectangle(geometry.CreatePoint2d(0, 0), geometry.CreatePoint2d(10, 10))
    dimension_1 = shape_1.Item(1)
    dimension_2 = shape_1.Item(4)
    
    #Add dimensional constraints
    sketch_1.DimensionConstraints.AddTwoPointDistance(dimension_1.StartSketchPoint, dimension_1.EndSketchPoint, 19201, geometry.CreatePoint2d(0, 0)) #kHorizontalDim
    sketch_1.DimensionConstraints.AddTwoPointDistance(dimension_2.StartSketchPoint, dimension_2.EndSketchPoint, 19202, geometry.CreatePoint2d(0, 0)) #kVerticalDim
    model_parameters.Item("d0").Value = user_parameters.Item("Length_1").Value
    model_parameters.Item("d1").Value = user_parameters.Item("Height_1").Value

    #Generate extrusion
    solid_profile_sketch_1 = sketch_1.Profiles.AddforSolid()
    extrude_1_definition = part_definition.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile_sketch_1, 20481) #kJoinOperation
    extrude_1_definition.SetDistanceExtent(extrusion_amount, 20994) #kNegativeExtentDirection
    part_definition.Features.ExtrudeFeatures.Add(extrude_1_definition)
    
    #Activate isometric view
    inventor.ActiveView.GoHome()
    inventor_part.Activate()

main() #Run the program