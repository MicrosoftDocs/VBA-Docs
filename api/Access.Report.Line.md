---
title: Report.Line method (Access)
keywords: vbaac10.chm13783
f1_keywords:
- vbaac10.chm13783
ms.prod: access
api_name:
- Access.Report.Line
ms.assetid: 9e640e37-c055-3dc3-b70e-0805cdc13561
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.Line method (Access)

The **Line** method draws lines and rectangles on a **Report** object when the **Print** event occurs.


## Syntax

_expression_.**Line** (_Step_ (_x1, y1_) - _Step_ (_x2, y2_), _Color_, _BF_)

_expression_ Required. A variable that represents a **[Report](Access.Report.md)** object. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Data type|Description|
|:---|:--------|:----------|
| _Step_ |_Keyword_ |Indicates that the starting point coordinates are relative to the current graphics position given by the current settings for the **[CurrentX](Access.Report.CurrentX.md)** and **[CurrentY](Access.Report.CurrentY.md)** properties of the _Object_ argument.|
|_x1, y1_ | **Single** |Indicates the coordinates of the starting point for the line or rectangle. The Scale properties (**[ScaleMode](Access.Report.ScaleMode.md)**, **ScaleLeft**, **ScaleTop**, **ScaleHeight**, and **ScaleWidth**) of the **Report** object specified by the _Object_ argument determine the unit of measure used. If this argument is omitted, the line begins at the position indicated by the **CurrentX** and **CurrentY** properties.|
|_x2, y2_ | **Single** |Required. Indicates the coordinates of the ending point for the line or rectangle. Ensure that the starting point and the ending point are separated by a hyphen ( - ).|
| _Color_ | **Long** |Indicates the RGB (red-green-blue) color used to draw the line. If this argument is omitted, the value of the **ForeColor** property is used. You can also use the **RGB** function or **QBColor** function to specify the color.|
| _B_ | |An option that creates a rectangle by using the coordinates as opposite corners of the rectangle.|
| _F_ | |_F_ cannot be used without _B_. If the _B_ option is used, the _F_ option specifies that the rectangle is filled with the same color used to draw the rectangle. If _B_ is used without _F_, the rectangle is filled with the color specified by the current settings of the **FillColor** and **BackStyle** properties. The default value for the **BackStyle** property is Normal for rectangles and lines.|


## Remarks

You can use this method only in an event procedure or a macro specified by the **OnPrint** or **OnFormat** event property for a report section, or the **OnPage** event property for a report.

To connect two drawing lines, make sure that one line begins at the end point of the previous line.

The width of the line drawn depends on the **[DrawWidth](Access.Report.DrawWidth.md)** property setting. The way a line or rectangle is drawn on the background depends on the settings of the **[DrawMode](Access.Report.DrawMode.md)** and **[DrawStyle](Access.Report.DrawStyle.md)** properties.

When you apply the **Line** method, the **CurrentX** and **CurrentY** properties are re-set to the end point specified by the _x2_ and _y2_ arguments.


## Example

The following example uses the **Line** method to draw a red rectangle five pixels inside the edge of a report named **EmployeeReport**. The **RGB** function is used to make the line red.

To try this example in Microsoft Access, create a new report named **EmployeeReport**. Paste the following code in the Declarations section of the report's module, and then switch to Print Preview.

```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
    ' Call the Drawline procedure 
    DrawLine 
End Sub 
 
Sub DrawLine() 
    Dim rpt As Report, lngColor As Long 
    Dim sngTop As Single, sngLeft As Single 
    Dim sngWidth As Single, sngHeight As Single 
 
    Set rpt = Reports!EmployeeReport 
    ' Set scale to pixels. 
    rpt.ScaleMode = 3 
    ' Top inside edge. 
    sngTop = rpt.ScaleTop + 5 
    ' Left inside edge. 
    sngLeft = rpt.ScaleLeft + 5 
    ' Width inside edge. 
    sngWidth = rpt.ScaleWidth - 10 
    ' Height inside edge. 
    sngHeight = rpt.ScaleHeight - 10 
    ' Make color red. 
    lngColor = RGB(255,0,0) 
    ' Draw line as a box. 
    rpt.Line(sngTop, sngLeft) - (sngWidth, sngHeight), lngColor, B 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
