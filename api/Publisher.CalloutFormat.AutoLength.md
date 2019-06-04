---
title: CalloutFormat.AutoLength property (Publisher)
keywords: vbapb10.chm2490627
f1_keywords:
- vbapb10.chm2490627
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.AutoLength
ms.assetid: ed874ec4-d4ce-5e3f-771a-8b3158f40707
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.AutoLength property (Publisher)

Returns an **[MsoTriState](Office.MsoTriState.md)** constant indicating whether the first segment of the callout line is scaled when the callout is moved. Applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour**). Read-only.


## Syntax

_expression_.**AutoLength**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Return value

MsoTriState


## Remarks

The **AutoLength** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library.

Use the **[AutomaticLength](Publisher.CalloutFormat.AutomaticLength.md)** method to set this property to **msoTrue**, and use the **[CustomLength](Publisher.CalloutFormat.CustomLength.md)** method to set this property to **msoFalse**.


## Example

This example switches between an automatically-scaling first segment and one with a fixed length for the callout line for the first shape in the publication. For the example to work, the shape must be a callout.

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