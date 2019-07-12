---
title: CalloutFormat.Length property (PowerPoint)
keywords: vbapp10.chm559014
f1_keywords:
- vbapp10.chm559014
ms.prod: powerpoint
api_name:
- PowerPoint.CalloutFormat.Length
ms.assetid: b0144e68-b495-0ef3-b228-599e56b7833e
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Length property (PowerPoint)

When the  **[AutoLength](PowerPoint.CalloutFormat.AutoLength.md)** property of the specified callout is set to **False**, the **Length** property returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only.


## Syntax

_expression_.**Length**

_expression_ A variable that represents a [CalloutFormat](PowerPoint.CalloutFormat.md) object.


## Remarks

Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour**). Use the **[CustomLength](PowerPoint.CalloutFormat.CustomLength.md)** method to set the value of this property for the **CalloutFormat** object.


## Example

If the first line segment in the callout named "co1" has a fixed length, this example specifies that the length of the first line segment in the callout named "co2" will also be fixed at that length. For the example to work, both callouts must have multiple-segment lines.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    With .Item("co1").Callout

        If Not .AutoLength Then len1 = .Length

    End With

    If len1 Then .Item("co2").Callout.CustomLength len1

End With
```


## See also


[CalloutFormat Object](PowerPoint.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]