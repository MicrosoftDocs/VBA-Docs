---
title: CalloutFormat.AutoLength property (Excel)
keywords: vbaxl10.chm104009
f1_keywords:
- vbaxl10.chm104009
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutoLength
ms.assetid: aadce7bf-e4b3-b56d-8a10-cf8183282149
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.AutoLength property (Excel)

Applies only to callouts whose lines consist of more than one segment (**[MsoCalloutType](office.msocallouttype.md)** types **msoCalloutThree** and **msoCalloutFour**). Read/write **[MsoTriState](Office.MsoTriState.md)**.


## Syntax

_expression_.**AutoLength**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

The first segment of the callout line (the segment attached to the text callout box) is scaled automatically whenever the callout is moved. 

The first segment of the callout retains the fixed length specified by the **[Length](excel.calloutformat.length.md)** property whenever the callout is moved.

This property is read-only. Use the **[AutomaticLength](Excel.CalloutFormat.AutomaticLength.md)** method to set this property to **msoTrue**, and use the **[CustomLength](Excel.CalloutFormat.CustomLength.md)** method to set this property to **msoFalse**.


## Example

This example toggles between an automatically scaling first segment and one with a fixed length for the callout line for shape one on _myDocument_. For the example to work, shape one must be a callout.

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]