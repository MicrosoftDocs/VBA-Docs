---
title: Shapes object (Project)
ms.prod: project-server
ms.assetid: 6e42040c-dd5a-de4c-afa8-f9e33d1e5054
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes object (Project)
Represents a collection of  **[Shape](./Project.shape.md)** objects in a custom report.

## Example

Use the  **[Report.Shapes](./Project.report.shapes.md)** property to get the **Shapes** collection object. In the following example, the report must be the active view to get the **Shapes** collection; otherwise, you get a run-time error 424 (Object required) in the `For Each oShape In oReport.Shapes` statement.


```vb
Sub ListShapesInReport()
    Dim oReports As Reports
    Dim oReport As Report
    Dim oShape As shape
    Dim reportName As String
    Dim msg As String
    Dim msgBoxTitle As String
    Dim numShapes As Integer
    
    numShapes = 0
    msg = ""
    reportName = "Table Tests"
    Set oReports = ActiveProject.Reports
    
    If oReports.IsPresent(reportName) Then
        ' Make the report the active view.
        oReports(reportName).Apply
        
        Set oReport = oReports(reportName)
        msgBoxTitle = "Shapes in report: '" & oReport.Name & "'"
    
        For Each oShape In oReport.Shapes
            numShapes = numShapes + 1
            msg = msg & numShapes & ". Shape type: " & CStr(oShape.Type) _
                & ", '" & oShape.Name & "'" & vbCrLf
        Next oShape
        
        If numShapes > 0 Then
            MsgBox Prompt:=msg, Title:=msgBoxTitle
        Else
            MsgBox Prompt:="This report contains no shapes.", _
                Title:=msgBoxTitle
        End If
    Else
         MsgBox Prompt:="The requested report, '" & reportName _
            & "', does not exist.", Title:="Report error"
    End If
End Sub
```


## Methods



|Name|
|:-----|
|[AddCallout](./Project.shapes.addcallout.md)|
|[AddChart](./Project.shapes.addchart.md)|
|[AddConnector](./Project.shapes.addconnector.md)|
|[AddCurve](./Project.shapes.addcurve.md)|
|[AddLabel](./Project.shapes.addlabel.md)|
|[AddLine](./Project.shapes.addline.md)|
|[AddPolyline](./Project.shapes.addpolyline.md)|
|[AddShape](./Project.shapes.addshape.md)|
|[AddTable](./Project.shapes.addtable.md)|
|[AddTextbox](./Project.shapes.addtextbox.md)|
|[AddTextEffect](./Project.shapes.addtexteffect.md)|
|[BuildFreeform](./Project.shapes.buildfreeform.md)|
|[Item](./Project.shapes.item.md)|
|[Range](./Project.shapes.range.md)|
|[SelectAll](./Project.shapes.selectall.md)|

## Properties



|Name|
|:-----|
|[Background](./Project.shapes.background.md)|
|[Count](./Project.shapes.count.md)|
|[Default](./Project.shapes.default.md)|
|[Parent](./Project.shapes.parent.md)|
|[Value](./Project.shapes.value.md)|

## See also


[Shape Object](./Project.shape.md)
[Report Object](./Project.report.md)
[ShapeRange Object](./Project.shaperange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]