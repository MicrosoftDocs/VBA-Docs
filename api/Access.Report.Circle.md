---
title: Report.Circle method (Access)
keywords: vbaac10.chm13782
f1_keywords:
- vbaac10.chm13782
ms.prod: access
api_name:
- Access.Report.Circle
ms.assetid: 4f5d24e2-75bf-3586-7e0d-0902adee61a6
ms.date: 03/09/2019
localization_priority: Normal
---


# Report.Circle method (Access)

The **Circle** method draws a circle, an ellipse, or an arc on a **Report** object when the **Print** event occurs.


## Syntax

_expression_.**Circle** (_Step_ (_x, y_), _Radius_, _Color_, _Start_, _End_, _Aspect_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Data type|Description|
|:---|:--------|:----------|
| _Step_ | **Keyword**|Indicates that the center of the circle, ellipse, or arc is relative to the current coordinates given by the current settings for the **[CurrentX](Access.Report.CurrentX.md)** and **[CurrentY](Access.Report.CurrentY.md)** properties of the _Object_ argument.|
| (_x, y_)| **Single**|Indicates the coordinates of the center point of the circle, ellipse, or arc. The Scale properties (**[ScaleMode](Access.Report.ScaleMode.md)**, **ScaleLeft**, **ScaleTop**, **ScaleHeight**, and **ScaleWidth**) of the **Report** object specified by the _Object_ argument determine the unit of measure used.|
| _Radius_ | **Single**|Indicates the radius of the circle, ellipse, or arc. The Scale properties (**ScaleMode**, **ScaleLeft**, **ScaleTop**, **ScaleHeight**, and **ScaleWidth**) of the **Report** object specified by the _Object_ argument determine the unit of measure used. By default, distances are measured in [twips](../language/glossary/vbe-glossary.md#twip).|
| _Color_ | **Long** |Indicates the RGB (red-green-blue) color of the circle outline. If this argument is omitted, the value of the **ForeColor** property is used. You can also use the **RGB** function or **QBColor** function to specify the color.|
| _Start_ | **Single**|When a partial circle or ellipse is drawn, the _Start_ argument specifies (in radians) the beginning position of the arc. The default value for the _Start_ argument is 0 radians. The range is -2 pi radians to 2 pi radians.|
| _End_ |**Single**|When a partial circle or ellipse is drawn, the _End_ argument specifies (in radians) the end position of the arc. The default value for the _End_ argument is 2 pi radians. The range is -2 pi radians to 2 pi radians.|
| _Aspect_ |**Single**| Indicates the aspect ratio of the circle. The default value is 1.0, which yields a perfect (nonelliptical) circle on any screen.|


## Remarks

You can use this method only in an event procedure or a macro specified by the event properties for a report section, or the **OnPage** event property for a report.

When drawing a partial circle or ellipse, if the _Start_ argument is negative, the **Circle** method draws a radius to the position specified by the _Start_ argument and treats the angle as positive. If the _End_ argument is negative, the **Circle** method draws a radius to the position specified by the _End_ argument and again treats the angle as positive. The **Circle** method always draws in a counterclockwise (positive) direction.

To fill a circle, set the **FillColor** and **FillStyle** properties of the report. Only a closed figure can be filled. Closed figures include circles, ellipses, and pie slices, which are arcs with radius lines drawn at both ends.

When drawing pie slices, if you need to draw a radius to angle 0 to form a horizontal line segment to the right, specify a very small negative value for the _Start_ argument rather than 0. For example, you might specify -.00000001 for the _Start_ argument.

You can omit an argument in the middle of the syntax, but you must include the argument's comma before including the next argument. If you omit a trailing argument, don't use any commas following the last argument that you specify.

The width of the line used to draw the circle, ellipse, or arc depends on the **[DrawWidth](Access.Report.DrawWidth.md)** property setting. The way the circle is drawn on the background depends on the settings of the **[DrawMode](Access.Report.DrawMode.md)** and **[DrawStyle](Access.Report.DrawStyle.md)** properties.

When you apply the **Circle** method, the **CurrentX** and **CurrentY** properties are set to the center point specified by the _x_ and _y_ arguments.


## Example

The following example uses the **Circle** method to draw a circle, and then create a pie slice within the circle and color it red.

To try this example in Microsoft Access, create a new report. Set the **OnPrint** property of the Detail section to [Event Procedure]. Enter the following code in the report's module, and then switch to Print Preview.

```vb
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
    Const conPI = 3.14159265359 
    Dim sngHCtr As Single, sngVCtr As Single 
    Dim sngRadius As Single 
    Dim sngStart As Single, sngEnd As Single 
 
    sngHCtr = Me.ScaleWidth / 2     ' Horizontal center. 
    sngVCtr = Me.ScaleHeight / 2     ' Vertical center. 
    sngRadius = Me.ScaleHeight / 3     ' Circle radius. 
    ' Draw circle. 
    Me.Circle(sngHCtr, sngVCtr), sngRadius 
    sngStart = -0.00000001             ' Start of pie slice. 
    sngEnd = -2 * conPI / 3             ' End of pie slice. 
    Me.FillColor = RGB(255,0,0)     ' Color pie slice red. 
    Me.FillStyle = 0                     ' Fill pie slice. 
    ' Draw pie slice within circle. 
    Me.Circle(sngHCtr, sngVCtr), sngRadius, , sngStart, sngEnd 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]