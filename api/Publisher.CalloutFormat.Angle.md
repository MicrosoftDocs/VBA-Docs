---
title: CalloutFormat.Angle property (Publisher)
keywords: vbapb10.chm2490625
f1_keywords:
- vbapb10.chm2490625
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Angle
ms.assetid: b65a1c87-db52-8703-135e-1fbb1efbeebe
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.Angle property (Publisher)

Returns or sets an **[MsoCalloutAngleType](office.msocalloutangletype.md)** constant that represents the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write.


## Syntax

_expression_.**Angle**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Remarks

If you set the value of this property to anything other than **msoCalloutAngleAutomatic**, the callout line maintains a fixed angle as you drag the callout.

## Example

This example sets the callout angle to 90 degrees for the first shape on the first page of the active publication. For this example to work, the specified shape must be a callout.

```vb
Sub SetCalloutAngle() 
 ActiveDocument.Pages(1).Shapes(1).Callout.Angle = msoCalloutAngle90 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]