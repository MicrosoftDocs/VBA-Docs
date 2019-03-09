---
title: Report.PSet method (Access)
keywords: vbaac10.chm13784
f1_keywords:
- vbaac10.chm13784
ms.prod: access
api_name:
- Access.Report.PSet
ms.assetid: 951a262b-b17b-9b95-b5f2-922d4aff9ce9
ms.date: 03/09/2019
localization_priority: Normal
---


# Report.PSet method (Access)

The **PSet** method sets a point on a **Report** object to a specified color when the **Print** event occurs.


## Syntax

_expression_.**PSet** (_Flags_, _x_, _y_, _Color_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Flags_|Required|**Integer**| A keyword that indicates that the coordinates are relative to the current graphics position given by the settings for the **[CurrentX](Access.Report.CurrentX.md)** and **[CurrentY](Access.Report.CurrentY.md)** properties of the _Object_ argument.|
| _x_|Required|**Single**|The horizontal coordinate of the point to set.|
| _y_|Required|**Single**|The vertical coordinate of the point to set.|
| _Color_|Required|**Long**|The RGB (red-green-blue) color to set the point to. If this argument is omitted, the value of the **ForeColor** property is used. You can also use the **RGB** function or **QBColor** function to specify the color.|

## Return value

Nothing


## Remarks

The size of the point depends on the **[DrawWidth](Access.Report.DrawWidth.md)** property setting. When the **DrawWidth** property is set to 1, the **PSet** method sets a single pixel to the specified color. When the **DrawWidth** property is greater than 1, the point is centered on the specified coordinates.

The way the point is drawn depends on the settings of the **[DrawMode](Access.Report.DrawMode.md)** and **[DrawStyle](Access.Report.DrawStyle.md)** properties.

When you apply the **PSet** method, the **CurrentX** and **CurrentY** properties are set to the point specified by the _x_ and _y_ arguments.

To clear a single pixel with the **PSet** method, specify the coordinates of the pixel and use &HFFFFFF (white) as the _Color_ argument.


## Example

The following example uses the **PSet** method to draw a line through the horizontal axis of a report.

To try this example in Microsoft Access, create a new report. Set the **OnPrint** property of the Detail section to [Event Procedure]. Enter the following code in the report's module, and then switch to Print Preview.

```vb
Sub Detail_Print(Cancel As Integer, PrintCount As Integer) 
 Dim sngMidPt As Single, intI As Integer 
 ' Set scale to pixels. 
 Me.ScaleMode = 3 
 ' Calculate midpoint. 
 sngMidPt = Me.ScaleHeight / 2 
 ' Loop to draw line down horizontal axis pixel by pixel. 
 For intI = 1 To Me.ScaleWidth 
 Me.PSet(intI, sngMidPt) 
 Next intI 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]