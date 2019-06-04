---
title: CalloutFormat.AutomaticLength method (Publisher)
keywords: vbapb10.chm2490384
f1_keywords:
- vbapb10.chm2490384
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.AutomaticLength
ms.assetid: 3772ad87-9808-5f25-0b9c-cdd7b1392ca1
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.AutomaticLength method (Publisher)

Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved.


## Syntax

_expression_.**AutomaticLength**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Remarks

Calling this method sets the **[AutoLength](Publisher.CalloutFormat.AutoLength.md)** property of the specified object to **msoTrue**.

Use the **[CustomLength](Publisher.CalloutFormat.CustomLength.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](Publisher.CalloutFormat.Length.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour**).


## Example

This example switches between an automatically-scaling first segment and one with a fixed length for the callout line for the first shape in the active publication. For the example to work, this shape must be a callout.

```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .AutoLength Then 
 .CustomLength Length:=50 
 Else 
 .AutomaticLength 
 End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]