# import python modules
import os
import math
import pandas
from win32com.client import Dispatch                         # Connecting to windows COM 
CATIA = Dispatch('CATIA.Application')                        # optional CATIA visibility
CATIA.Visible = True                                         # Create an empty part
partDocument1 = CATIA.Documents.Add("Part")
df = pandas.read_excel('BS_UB.xls', sheetname='Dimensions')  # Input excel file 

B=df.iloc[41,2]                                              # Cell location of Flange Length
D=df.iloc[41,1]                                              # Cell location of Web Length
T=df.iloc[41,4]                                              # Cell location of Flange Thickness
t=df.iloc[41,3]                                              # Cell location of Web Thickness
r=df.iloc[41,5]                                              # Cell location of Fillet Radius

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
NewLine1 = NewSketch.Factory2D.CreateLine((-B/2), (D/2)  , (B/2), (D/2))
NewLine2 = NewSketch.Factory2D.CreateLine((B/2), (D/2) , (B/2), (D/2-T))
NewLine3 = NewSketch.Factory2D.CreateLine((B/2), (D/2-T) , (t/2+r), (D/2-T))
NewLine4 = NewSketch.Factory2D.CreateLine((t/2), (D/2-T-r) , (t/2), (-D/2+T+r))
NewLine5 = NewSketch.Factory2D.CreateLine((B/2), (-D/2+T) , (t/2+r), (-D/2+T))
NewLine6 = NewSketch.Factory2D.CreateLine((B/2), (-D/2) , (B/2), (-D/2+T))
NewLine7 = NewSketch.Factory2D.CreateLine((-B/2), (-D/2)  , (B/2), (-D/2))
NewLine8 = NewSketch.Factory2D.CreateLine((-B/2), (-D/2) , (-B/2), (-D/2+T))
NewLine9 = NewSketch.Factory2D.CreateLine((-B/2), (-D/2+T) , (-t/2-r), (-D/2+T))
NewLine10 = NewSketch.Factory2D.CreateLine((-t/2), (-D/2+T+r) , (-t/2), (D/2-T-r))
NewLine11 = NewSketch.Factory2D.CreateLine((-B/2), (D/2-T) , (-t/2-r), (D/2-T))
NewLine12 = NewSketch.Factory2D.CreateLine((-B/2), (D/2) , (-B/2), (D/2-T))

NewArc1 = NewSketch.Factory2D.CreateCircle((t/2+r), (D/2-T-r), r, math.radians(90), math.radians(180))
NewArc2 = NewSketch.Factory2D.CreateCircle((t/2+r), (-D/2+T+r), r, math.radians(180), math.radians(270))
NewArc3 = NewSketch.Factory2D.CreateCircle((-t/2-r), (-D/2+T+r), r, math.radians(270), 0)
NewArc4 = NewSketch.Factory2D.CreateCircle((-t/2-r), (D/2-T-r), r, 0, math.radians(90))

partDocument1.Part.Update()
NewSketch.CloseEdition()
print(B,D,T,t,r)
