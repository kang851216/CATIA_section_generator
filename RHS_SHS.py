# import python modules
import os
import math
from win32com.client import Dispatch            # Connecting to windows COM 
CATIA = Dispatch('CATIA.Application')           # optional CATIA visibility
CATIA.Visible = True                            # Create an empty part
partDocument1 = CATIA.Documents.Add("Part")

B=100                                           # Width
D=100                                           # Height
t=4                                             # Thickness
r=4                                             # Root radius
b=B-t                                           # Inside width
d=D-t                                           # Inside height

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
#Point1 = NewSketch.Factory2D.CreatePoint((-B/2+r), (D/2))
#Point2 = NewSketch.Factory2D.CreatePoint((B/2-r), (D/2))
#Point3 = NewSketch.Factory2D.CreatePoint((B/2), (D/2-r))
#Point4 = NewSketch.Factory2D.CreatePoint((B/2), (-D/2+r))
#Point5 = NewSketch.Factory2D.CreatePoint((B/2-r), (-D/2))
#Point6 = NewSketch.Factory2D.CreatePoint((-B/2+r), (-D/2))
#Point7 = NewSketch.Factory2D.CreatePoint((-B/2), (-D/2+r))
#Point8 = NewSketch.Factory2D.CreatePoint((-B/2), (D/2-r))
#CirclePoint1 = NewSketch.Factory2D.CreatePoint((-B/2+r), (D/2-r))
#CirclePoint2 = NewSketch.Factory2D.CreatePoint((B/2-r), (D/2-r))
#CirclePoint3 = NewSketch.Factory2D.CreatePoint((B/2-r), (-D/2+r))
#CirclePoint4 = NewSketch.Factory2D.CreatePoint((-B/2+r), (-D/2+r))

NewLine1 = NewSketch.Factory2D.CreateLine((-B/2+r), (D/2)  , (B/2-r), (D/2))
NewLine2 = NewSketch.Factory2D.CreateLine((B/2), (D/2-r) , (B/2), (-D/2+r))
NewLine3 = NewSketch.Factory2D.CreateLine((B/2-r), (-D/2) , (-B/2+r), (-D/2))
NewLine4 = NewSketch.Factory2D.CreateLine((-B/2), (-D/2+r) , (-B/2), (D/2-r))
NewArc1 = NewSketch.Factory2D.CreateCircle((-B/2+r), (D/2-r), r, math.radians(90), math.radians(180))
NewArc2 = NewSketch.Factory2D.CreateCircle((B/2-r), (D/2-r), r, 0, math.radians(90))
NewArc3 = NewSketch.Factory2D.CreateCircle((B/2-r), (-D/2+r), r, math.radians(270), 0)
NewArc4 = NewSketch.Factory2D.CreateCircle((-B/2+r), (-D/2+r), r, math.radians(180), math.radians(270))

NewLine5 = NewSketch.Factory2D.CreateLine((-b/2+r), (d/2)  , (b/2-r), (d/2))
NewLine6 = NewSketch.Factory2D.CreateLine((b/2), (d/2-r) , (b/2), (-d/2+r))
NewLine7 = NewSketch.Factory2D.CreateLine((b/2-r), (-d/2) , (-b/2+r), (-d/2))
NewLine8 = NewSketch.Factory2D.CreateLine((-b/2), (-d/2+r) , (-b/2), (d/2-r))
NewArc5 = NewSketch.Factory2D.CreateCircle((-b/2+r), (d/2-r), r, 1.57079, 3.14159)
NewArc6 = NewSketch.Factory2D.CreateCircle((b/2-r), (d/2-r), r, 0, 1.57079)
NewArc7 = NewSketch.Factory2D.CreateCircle((b/2-r), (-d/2+r), r, 4.71238, 0)
NewArc8 = NewSketch.Factory2D.CreateCircle((-b/2+r), (-d/2+r), r, 3.14159, 4.71238)

partDocument1.Part.Update()
NewSketch.CloseEdition()
