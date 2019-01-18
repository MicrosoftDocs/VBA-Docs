---
title: CalloutFormat object (Excel)
keywords: vbaxl10.chm104000
f1_keywords:
- vbaxl10.chm104000
ms.prod: excel
api_name:
- Excel.CalloutFormat
ms.assetid: d9d7d279-04ef-dbee-23cd-ddd606ed917d
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat object (Excel)

Contains properties and methods that apply to line callouts.


## Remarks

Use the  **[Callout](Excel.Shape.Callout.md)** property to return a **CalloutFormat** object.


## Example

 The following example specifies the following attributes of shape three (a line callout) on _myDocument_ : the callout will have a vertical accent bar that separates the text from the callout line; the angle between the callout line and the side of the callout text box will be 30 degrees; there will be no border around the callout text; the callout line will be attached to the top of the callout text box; and the callout line will contain two segments. For this example to work, shape three must be a callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Callout 
 .Accent = True 
 .Angle = msoCalloutAngle30 
 .Border = False 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
End With
```


## See also


[Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]