---
title: CalloutFormat.AutoLength property (Excel)
keywords: vbaxl10.chm104009
f1_keywords:
- vbaxl10.chm104009
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutoLength
ms.assetid: aadce7bf-e4b3-b56d-8a10-cf8183282149
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.AutoLength property (Excel)

Applies only to callouts whose lines consist of more than one segment (types  **msoCalloutThree** and **msoCalloutFour**). Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_. `AutoLength`

_expression_ A variable that represents a [CalloutFormat](Excel.CalloutFormat.md) object.


## Remarks



| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse**. The first segment of the callout retains the fixed length specified by the **Length** property whenever the callout is moved.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue**. The first segment of the callout line (the segment attached to the text callout box) is scaled automatically whenever the callout is moved.|

This property is read-only. Use the  **[AutomaticLength](Excel.CalloutFormat.AutomaticLength.md)** method to set this property to **msoTrue**, and use the **[CustomLength](Excel.CalloutFormat.CustomLength.md)** method to set this property to **mosFalse**.


## Example

This example toggles between an automatically scaling first segment and one with a fixed length for the callout line for shape one on  `myDocument`. For the example to work, shape one must be a callout.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Callout 
    If .AutoLength Then 
        .CustomLength 50 
    Else 
        .AutomaticLength 
    End If 
End With
```


## See also


[CalloutFormat Object](Excel.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]