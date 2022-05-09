Set MyDocument = Worksheets(1)
With MyDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200)
.AddNodes msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300
.AddNodes msoSegmentCurve, msoEditingAuto, 480, 200
.AddNodes msoSegmentLine, msoEditingAuto, 480, 400
.AddNodes msoSegmentLine, msoEditingAuto, 360, 200
.ConvertToShape
End With
