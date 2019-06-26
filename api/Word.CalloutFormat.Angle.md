---
title: CalloutFormat.Angle property (Word)
keywords: vbawd10.chm163905637
f1_keywords:
- vbawd10.chm163905637
ms.prod: word
api_name:
- Word.CalloutFormat.Angle
ms.assetid: b5178aa0-c2e3-dc59-766d-7ce5b2e7c762
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Angle property (Word)

Returns or sets the angle of the callout line. Read/write  **[MsoCalloutAngleType](Office.MsoCalloutAngleType.md)**.


## Syntax

_expression_.**Angle**

_expression_ A variable that represents a '[CalloutFormat](Word.CalloutFormat.md)' object.


## Remarks

If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. If you set the value of this property to anything other than  **msoCalloutAngleAutomatic**, the callout line maintains a fixed angle as you drag the callout.


> [!NOTE] 
> Setting this property to  **msoCalloutAngleMixed** will cause an error. **msoCalloutAngleMixed** is a return value only. It indicates a combination of the other states.


## Example

This example sets the callout angle to 90 degrees for a callout named "co1" on the active document.


```vb
ActiveDocument.Shapes("co1").Callout.Angle = msoCalloutAngle90
```


## See also


[CalloutFormat Object](Word.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]