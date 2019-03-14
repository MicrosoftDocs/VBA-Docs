---
title: Report.CurrentY property (Access)
keywords: vbaac10.chm13742
f1_keywords:
- vbaac10.chm13742
ms.prod: access
api_name:
- Access.Report.CurrentY
ms.assetid: 040c0b5d-f7d6-2fa1-e34d-f69799f0b273
ms.date: 03/15/2019
localization_priority: Normal
---


# Report.CurrentY property (Access)

You can use the **CurrentY** property (along with the **CurrentX** property) to specify the horizontal and vertical coordinates for the starting position of the next printing and drawing method on a report. Read/write **Single**.


## Syntax

_expression_.**CurrentY**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

For example, you can use these properties to determine where the center point of a circle is drawn on a report section.

The coordinates are measured from the upper-left corner of the report section that contains the reference to the **CurrentX** or **CurrentY** property. The **CurrentX** property setting is 0 at the section's left edge, and the **CurrentY** property setting is 0 at its top edge.

You can set the **CurrentX** and **CurrentY** properties in an event procedure specified by the **[OnPrint](Access.Section.OnPrint.md)** property setting of a report section.

Use the **[ScaleMode](Access.Report.ScaleMode.md)** property to define the unit of measure, such as points, pixels, characters, inches, millimeters, or centimeters, that the coordinates will be based on.

When you use the following graphics methods, the **CurrentX** and **CurrentY** property settings are changed as indicated.

|Method|Sets CurrentX, CurrentY properties to|
|:-----|:-----|
|**[Circle](Access.Report.Circle.md)**|The center of the object.|
|**[Line](Access.Report.Line.md)**|The end point of the line (the _x2_, _y2_ coordinates).|
|**[Print](Access.Report.Print.md)**|The next print position.|

## Example

The following example uses the **Print** method to display text on a report named **Report1**. It uses the **TextWidth** and **TextHeight** methods to center the text vertically and horizontally.

```vb
Private Sub Detail_Format(Cancel As Integer, _ 
 FormatCount As Integer) 
 Dim rpt as Report 
 Dim strMessage As String 
 Dim intHorSize As Integer, intVerSize As Integer 
 
 Set rpt = Me 
 strMessage = "DisplayMessage" 
 With rpt 
 'Set scale to pixels, and set FontName and 
 'FontSize properties. 
 .ScaleMode = 3 
 .FontName = "Courier" 
 .FontSize = 24 
 End With 
 ' Horizontal width. 
 intHorSize = Rpt.TextWidth(strMessage) 
 ' Vertical height. 
 intVerSize = Rpt.TextHeight(strMessage) 
 ' Calculate location of text to be displayed. 
 Rpt.CurrentX = (Rpt.ScaleWidth/2) - (intHorSize/2) 
 Rpt.CurrentY = (Rpt.ScaleHeight/2) - (intVerSize/2) 
 ' Print text on Report object. 
 Rpt.Print strMessage 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]