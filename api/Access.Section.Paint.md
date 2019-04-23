---
title: Section.Paint event (Access)
keywords: vbaac10.chm14238
f1_keywords:
- vbaac10.chm14238
ms.prod: access
api_name:
- Access.Section.Paint
ms.assetid: f68d981d-8371-cf0d-9da4-063aaa0f0907
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.Paint event (Access)

Occurs when the specified section is redrawn.


## Syntax

_expression_.**Paint**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Example

The following example shows how to set the **BackColor** property of a control based on its value.

```vb
Private Sub SetControlFormatting()
    If (Me.AvgOfRating >= 8) Then
        Me.AvgOfRating.BackColor = vbGreen
    ElseIf (Me.AvgOfRating >= 5) Then
        Me.AvgOfRating.BackColor = vbYellow
    Else
        Me.AvgOfRating.BackColor = vbRed
    End If
End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' size the width of the rectangle
    Dim lngOffset As Long
    lngOffset = (Me.boxInside.Left - Me.boxOutside.Left) * 2
    Me.boxInside.Width = (Me.boxOutside.Width * (Me.AvgOfRating / 10)) - lngOffset
    
    ' do conditional formatting for the control in print preview
    SetControlFormatting
End Sub

Private Sub Detail_Paint()
    ' do conditional formatting for the control in report view
    SetControlFormatting
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]