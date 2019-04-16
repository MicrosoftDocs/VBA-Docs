---
title: CalloutFormat.Angle property (Excel)
keywords: vbaxl10.chm104007
f1_keywords:
- vbaxl10.chm104007
ms.prod: excel
api_name:
- Excel.CalloutFormat.Angle
ms.assetid: 8f3dab54-4597-e22c-ae3e-cf894849b668
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.Angle property (Excel)

Returns or sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write **[MsoCalloutAngleType](Office.MsoCalloutAngleType.md)**.


## Syntax

_expression_.**Angle**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

If you set the value of this property to anything other than **msoCalloutAngleAutomatic**, the callout line maintains a fixed angle as you drag the callout.


## Example

This example sets to 90 degrees the callout angle for a callout named callout1 on _myDocument_.

```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes("callout1").Callout.Angle = msoCalloutAngle90
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]