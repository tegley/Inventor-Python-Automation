#Import statements
import win32com.client as wc
import os
import time

def main(): #Control Inventor and execution of functions
    global inventor
    os.startfile(r'C:\Program Files\Autodesk\Inventor 2023\Bin\Inventor.exe') #Open Inventor from File Path
    print("Inventor has opened - please wait.")
    time.sleep(30)
    inventor = wc.GetActiveObject('Inventor.Application') #Connect Inventor to Python
    inventor_part = inventor.Documents.Add(12290, inventor.FileManager.GetTemplateFile(12290,8963))

    
    radius, height, extrusion_amount, radius_choice, height_choice, extrusion_amount_choice = Collect_User_Data()
    part_definition, user_parameters, model_parameters, reference_parameters, geometry, sketch_1, sketch_2, origin, dim_number = Collect_Inventor_Data()

    #Execute vessel saddle creation
    line_2, line_3 = Vessel_Saddle_Sketch_1(user_parameters, model_parameters, sketch_1, geometry, radius_choice, origin, dim_number)   
    Vessel_Saddle_Sketch_2(user_parameters, model_parameters, sketch_2, geometry, origin, dim_number, radius, height_choice, height, line_2, line_3)
    Vessel_Saddle_Extrusion(inventor_part, part_definition, sketch_2, extrusion_amount_choice)
    
    #Reactivate Part Document and Turn Off Sketch Dimensions
    inventor_part.Activate()
    inventor_part.ObjectVisibility.SketchDimensions = False

    print("Operation complete!")

def Collect_User_Data(): #Create a part object according to user input
    radius = float(input("Select radius of the vessel -> "))
    radius_choice = str(radius)
    radius_choice += " in"
    height = float(input("Select height of the vessel saddle -> "))
    height_choice = str(height)
    height_choice += " in"
    extrusion_amount = float(input("Select extrusion amount -> "))
    extrusion_amount_choice = str(extrusion_amount)
    extrusion_amount_choice += " in"

    return(radius, height, extrusion_amount, radius_choice, height_choice, extrusion_amount_choice)

def Collect_Inventor_Data(inventor_part):
    #Define the part file and corresponding parameters
    part_definition = inventor_part.ComponentDefinition
    user_parameters = part_definition.Parameters.UserParameters
    model_parameters = part_definition.Parameters.ModelParameters
    reference_parameters = part_definition.Parameters.ReferenceParameters
    geometry = inventor.TransientGeometry

    #Create new sketches
    sketch_1 = part_definition.Sketches.Add(part_definition.WorkPlanes.Item(3))
    sketch_2 = part_definition.Sketches.Add(part_definition.WorkPlanes.Item(3))

    #Define the origin
    origin = sketch_1.AddByProjectingEntity(part_definition.WorkPoints.Item(1))

    #Define the dimension number
    dim_number = [-1]
    return(part_definition, user_parameters, model_parameters, reference_parameters, geometry, sketch_1, sketch_2, origin, dim_number)


#Create a vessel saddle
def Vessel_Saddle_Sketch_1(user_parameters, model_parameters, sketch_1, geometry, radius_choice, origin, dim_number): #Create the general vessel shape and construction lines
    
    #Create circle with user-defined radius and constrain it to the center
    user_parameters.AddByExpression("Radius_1", radius_choice, 11272) #kInchLengthUnits
    shape_1 = sketch_1.SketchCircles.AddByCenterRadius(origin, 5)
    sketch_1.DimensionConstraints.AddRadius(shape_1, geometry.CreatePoint2d(0, 0))    
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Radius_1").Value    
    sketch_1.GeometricConstraints.AddCoincident(shape_1.CenterSketchPoint, origin)

    #Creating points for use in line creation
    start = geometry.CreatePoint2d(0, 0)
    end  = geometry.CreatePoint2d(0, -10)

    #Create first/centered construction line    
    line_1 = sketch_1.SketchLines.AddByTwoPoints(start, end) 
    sketch_1.DimensionConstraints.AddTwoPointDistance(line_1.StartSketchPoint, line_1.EndSketchPoint, 19203, geometry.CreatePoint2d(0, 0)) #kAlignedDim
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Radius_1").Value
    sketch_1.GeometricConstraints.AddCoincident(line_1.StartSketchPoint, origin)
    sketch_1.GeometricConstraints.AddVertical(line_1)

    #Create second/left construction line
    line_2 = sketch_1.SketchLines.AddByTwoPoints(start, end) 
    sketch_1.DimensionConstraints.AddTwoPointDistance(line_2.StartSketchPoint, line_2.EndSketchPoint, 19203, geometry.CreatePoint2d(0, 0)) #kAlignedDim
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Radius_1").Value

    #Add angle constraint
    sketch_1.DimensionConstraints.AddTwoLineAngle(line_2, line_1, geometry.CreatePoint2d(-5,-5))
    user_parameters.AddByExpression("Angle_1", "60 deg", 11279) #kDegreeAngleUnits
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Angle_1").Value
    
    #Align line to the center
    sketch_1.GeometricConstraints.AddVerticalAlign(line_2.StartSketchPoint, origin)
    sketch_1.GeometricConstraints.AddHorizontalAlign(line_2.StartSketchPoint, origin)

    #Create third/right construction line
    line_3 = sketch_1.SketchLines.AddByTwoPoints(start, end) 
    sketch_1.DimensionConstraints.AddTwoPointDistance(line_3.StartSketchPoint, line_3.EndSketchPoint, 19203, geometry.CreatePoint2d(0, 0)) #kAlignedDim
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Radius_1").Value
    
    #Add angle constraint
    sketch_1.DimensionConstraints.AddTwoLineAngle(line_1, line_3, geometry.CreatePoint2d(5,-5))
    user_parameters.AddByExpression("Angle_2", "60 deg", 11279) #kDegreeAngleUnits
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Angle_2").Value

    #Align line to the center
    sketch_1.GeometricConstraints.AddVerticalAlign(line_3.StartSketchPoint, origin)
    sketch_1.GeometricConstraints.AddHorizontalAlign(line_3.StartSketchPoint, origin)

    #Update document with changes    
    inventor.ActiveDocument.Update()
    return(line_2, line_3)    

def Vessel_Saddle_Sketch_2(user_parameters, model_parameters, sketch_2, geometry, origin, dim_number, radius, height_choice, height, line_2, line_3): #Create saddle shape and connect it
    
    #Project line start points
    height_1_start = sketch_2.AddByProjectingEntity(line_2.EndSketchPoint)
    height_2_start = sketch_2.AddByProjectingEntity(line_3.EndSketchPoint)
    
    #Create height lines
    x_cord = float(line_2.EndSketchPoint.Geometry.Y)
    point_1_1 = float(line_2.EndSketchPoint.Geometry.X)
    point_1_2 = float(line_2.EndSketchPoint.Geometry.X) * -1
    point_2 = x_cord - height
    height_1 = sketch_2.SketchLines.AddByTwoPoints(height_1_start, geometry.CreatePoint2d(point_1_1, point_2)) 
    height_2 = sketch_2.SketchLines.AddByTwoPoints(height_2_start, geometry.CreatePoint2d(point_1_2, point_2))

    #Straighten out height lines
    sketch_2.GeometricConstraints.AddVertical(height_1)
    sketch_2.GeometricConstraints.AddVertical(height_2)

    #Constrain height lines
    sketch_2.GeometricConstraints.AddEqualLength(height_1, height_2)

    #Add length of vessel saddle
    saddle_width = sketch_2.SketchLines.AddByTwoPoints(height_1.EndSketchPoint, height_2.EndSketchPoint)
    sketch_2.DimensionConstraints.AddTwoPointDistance(height_1.EndSketchPoint, height_2.EndSketchPoint, 19203, geometry.CreatePoint2d(0, 0), True) #kAlignedDim  
    dim_number[0] +=1

    #Add height of vessel saddle (distance from origin to the bottom of the saddle)
    sketch_2.DimensionConstraints.AddTwoPointDistance(height_1.EndSketchPoint, origin, 19202, geometry.CreatePoint2d(point_1_1, point_2)) #kVerticalDim 
    user_parameters.AddByExpression("Height_1", height_choice, 11272) #kInchLengthUnits
    model_parameters.Item(Dimension_Automation(dim_number)).Value = user_parameters.Item("Height_1").Value

    #Add arc connecting the start of the height lines
    math = radius*-2.54
    midpoint = geometry.CreatePoint2d(0,math)
    arc = sketch_2.SketchArcs.AddByThreePoints(height_1_start, midpoint, height_2_start)

    #Update part with changes
    inventor.ActiveDocument.Update()

def Vessel_Saddle_Extrusion(inventor_part, part_definition, sketch_2, extrusion_amount_choice): #Extrude vessel saddle by appropriate amount of inches
    #Reactivate part view
    inventor_part.Activate()
    
    #Create extrusion profile and definition
    solid_profile_sketch_2 = sketch_2.Profiles.AddforSolid()
    extrude_1_definition = part_definition.Features.ExtrudeFeatures.CreateExtrudeDefinition(solid_profile_sketch_2, 20481)
    extrude_1_definition.SetDistanceExtent(extrusion_amount_choice, 20993)

    #Add extrusion
    part_definition.Features.ExtrudeFeatures.Add(extrude_1_definition)
    
    #Update part with changes
    inventor.ActiveDocument.Update()

#Secondary functions
def Dimension_Automation(dim_number): #Keep track of the dimensions being created by automatically incrementing the dimension number
    dim_number[0] +=1
    notation = 'd' + str(dim_number[0])
    return(notation)

main()
