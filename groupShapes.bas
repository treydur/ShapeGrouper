Attribute VB_Name = "Module1"
'Need to reference excel regular expression 5.5 in vba references

Sub groupShapes()

Dim regEx As New RegExp
Dim strPattern As String: strPattern = "^Rectangle ([0-9]|[0-9][0-9]|[0-9][0-9][0-9])$"

For j = 1 To 34
    Set myDocument = ActivePresentation.Slides(j)

    'For Each s In myDocument.Shapes
    'For some reason you have to iterate backwards when looping through a collection that changes
    'The collection changes because groups are being added ~ shapes being removed
    For intLoop = myDocument.Shapes.Count To 1 Step -1
        Set s = myDocument.Shapes(intLoop)
        With regEx
            .Pattern = strPattern
        End With
        ' Debug.Print (regEx.Test(s.Name) & " " & s.Name)
        If regEx.Test(s.Name) Then
            'Adds an in visible 16 point star on top of the rectangle
            With myDocument.Shapes.AddShape(msoShape16pointStar, s.Left - 7, s.Top - 8, s.Width + 14, s.Height + 16)
                .Name = s.Name & " outline"
                newShapeName = .Name
                .Fill.Visible = False
                .Line.Visible = False
            End With
            Debug.Print ("added shape")
            
            myDocument.Shapes.Range(Array(newShapeName, s.Name)).Group
        End If
    Next
Next
 
End Sub
