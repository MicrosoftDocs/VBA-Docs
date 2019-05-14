---
title: Shape.TextEffect property (Project)
ms.prod: project-server
ms.assetid: 12fa0951-e3a5-807e-bebb-bff82650d200
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.TextEffect property (Project)
Gets text formatting properties for the shape. Read-only  **[TextEffectFormat](https://msdn.microsoft.com/library/office/ff834714%28v=office.15%29)**.

## Syntax

_expression_.**TextEffect**

_expression_ A variable that represents a **[Shape](Project.Shape.md)** object.


## Example

The following example sets the foreground color of text in a text frame to red, the foreground color of the text box shape to a yellowish tan, and then uses the  **TextEffect** property to set font properties.


```vb
Sub FormatTextBox()
    Dim theReport As Report
    Dim textShape As shape
    Dim reportName As String
    
    reportName = "Textbox report"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    Set textShape = theReport.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 50, 350, 80)
    
    textShape.TextFrame2.TextRange.Text = "This is a test. It is only a test. "
    textShape.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.RGB = &H2020CC
    textShape.Fill.ForeColor.RGB = &H88CCCC
    
    With textShape.TextEffect
        .FontName = "Courier New"
        .FontBold = True
        .FontItalic = True
        .FontSize = 28
    End With
End Sub
```


## Property value

 **TEXTEFFECTFORMAT**


## See also


[Shape Object](Project.shape.md)
[ShapeRange.TextEffect Property](Project.shaperange.texteffect.md)
[TextEffectFormat](https://msdn.microsoft.com/library/office/ff834714%28v=office.15%29)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]