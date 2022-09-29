# import python modules
import os
import math
from win32com.client import Dispatch            # Connecting to windows COM 
CATIA = Dispatch('CATIA.Application')           # optional CATIA visibility
CATIA.Visible = True                            # Create an empty part
partDocument1 = CATIA.Documents.Add("Part")

B=100                                                               # Flange width
D=430                                                               # Web height
T=19                                                                # Flange thickness
t=11                                                                # Web thickness
r=15                                                                # Root radius
cal_coe=0.982+(0.003*(r-9))                                         # Calibration coefficient 
b_y= ((2*T*B*B/2)+(t*(D-2*T)*t/2))/((2*T*B)+(t*(D-2*T)))*cal_coe    # Center of area


# Create the point that the plane and sketch reference
Xcoord=0
Ycoord=0
Zcoord=0
NewPoint = CATIA.ActiveDocument.Part.HybridShapeFactory.AddNewPointCoord(Xcoord, Ycoord, Zcoord)
Mainbody = CATIA.ActiveDocument.Part.MainBody
Mainbody.InsertHybridShape(NewPoint)

# Create reference plane for sketch
AxisXY = CATIA.ActiveDocument.Part.OriginElements.PlaneXY
Referenceplane = CATIA.ActiveDocument.Part.CreateReferenceFromObject(AxisXY)
Referencepoint = CATIA.ActiveDocument.Part.CreateReferenceFromObject(NewPoint)
NewPlane = CATIA.ActiveDocument.Part.HybridShapeFactory.AddNewPlaneOffsetPt(Referenceplane, Referencepoint)
Mainbody.InsertHybridShape(NewPlane)

# create sketch
sketches1 = CATIA.ActiveDocument.Part.Bodies.Item("PartBody").Sketches
reference1 = partDocument1.part.OriginElements.PlaneXY
NewSketch = sketches1.Add(reference1)
CATIA.ActiveDocument.Part.InWorkObject = NewSketch

#Drawing the sketch
NewSketch.OpenEdition()
#                                  Start (H,   V)  End(H,V)
NewLine1 = NewSketch.Factory2D.CreateLine((-b_y), (D/2)  , (B-b_y), (D/2))
NewLine2 = NewSketch.Factory2D.CreateLine((B-b_y), (D/2) , (B-b_y), (D/2-T))
NewLine3 = NewSketch.Factory2D.CreateLine((B-b_y), (D/2-T) , (-b_y+t+r), (D/2-T))
NewLine4 = NewSketch.Factory2D.CreateLine((-b_y+t), (D/2-T-r) , (-b_y+t), (-D/2+T+r))
NewLine5 = NewSketch.Factory2D.CreateLine((B-b_y), (-D/2+T) , (-b_y+t+r), (-D/2+T))
NewLine6 = NewSketch.Factory2D.CreateLine((B-b_y), (-D/2) , (B-b_y), (-D/2+T))
NewLine7 = NewSketch.Factory2D.CreateLine((B-b_y), (-D/2)  , (-b_y), (-D/2))
NewLine8 = NewSketch.Factory2D.CreateLine((-b_y), (-D/2) , (-b_y), (D/2))

NewArc1 = NewSketch.Factory2D.CreateCircle((-b_y+t+r), (D/2-T-r), r, math.radians(90), math.radians(180))
NewArc2 = NewSketch.Factory2D.CreateCircle((-b_y+t+r), (-D/2+T+r), r, math.radians(180), math.radians(270))

partDocument1.Part.Update()
NewSketch.CloseEdition()

