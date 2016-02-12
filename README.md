# test3
trying to get VBA code as a repository
Sub power_export()
    Application.ScreenUpdating = False
    Dim oWSS As Object
    Dim oPPTApp As Object
    Dim oPPTPres As Object
    Dim oPres As Object
    Dim oSlide As Object
    Dim oShape As Object
    Dim strName, strPeriod, filName, strQuant As String
    Dim intCol, intStart, intRow, intEndrow, intFlag As Integer, lstRow As Integer, lstCol As Integer
    
    Set oWSS = CreateObject("WScript.Shell")
    filName = oWSS.specialfolders("Desktop") & "\J&J Market Mirror Reporting Tool Charts.pptx"
    intFlag = 0
    Call invisible
    
    If Dir(filName) = "" Then
        Set oPPTApp = CreateObject("PowerPoint.Application")
        intFlag = 1
        oPPTApp.visible = True
        Set oPPTPres = oPPTApp.presentations.Add(msoTrue)
    Else
        Set oPPTApp = GetObject(filName).Application
        oPPTApp.visible = True
        For Each oPres In oPPTApp.presentations
            If oPres.Name = Right(filName, Len(filName) - InStrRev(filName, "\")) Then
                oPres.Windows(1).Activate
                Set oPPTPres = oPPTApp.activepresentation
                Exit For
            End If
        Next oPres
    End If
    
    If oPPTPres.Slides.count = 0 Then
        Set oSlide = oPPTPres.Slides.Add(1, 15)
    Else
        Set oSlide = oPPTPres.Slides.Add(oPPTPres.Slides.count + 1, 15)
        oPPTPres.Slides(oPPTPres.Slides.count).Select
    End If
    
    strQuant = Sheets("uRanges").Range("k2").Value
    
    lstRow = ActiveSheet.Cells(ActiveSheet.Rows.count, "D").End(xlUp).Row
    lstCol = ActiveSheet.Cells(10, ActiveSheet.Columns.count).End(xlToLeft).Column

    'This first part copies the data
        ActiveSheet.Range(Cells(1, 1), Cells(lstRow + 1, lstCol + 3)).Copy
        'And this part pastes the data into PowerPoint
        With oPPTPres
            oSlide.Shapes.placeholders(1).Select
            .Windows(1).View.PasteSpecial 2
            With Selection
                oPPTApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                oPPTApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True
            End With
        End With
    
    oPPTApp.WindowState = 2
    Application.CutCopyMode = False
    If intFlag = 1 Then
        oPPTPres.SaveAs Filename:=filName
    Else
        oPPTPres.Save
    End If
    Set oPPTPres = Nothing
    Set oPPTApp = Nothing
    Call visible
    Application.ScreenUpdating = True
End Sub
