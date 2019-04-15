---
title: CalloutFormat.AutomaticLength method (Excel)
keywords: vbaxl10.chm104002
f1_keywords:
- vbaxl10.chm104002
ms.prod: excel
api_name:
- Excel.CalloutFormat.AutomaticLength
ms.assetid: e82093e0-7b84-c2c8-8365-6fe05298d55b
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.AutomaticLength method (Excel)

Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. 

Use the **[CustomLength](Excel.CalloutFormat.CustomLength.md)** method to specify that the first segment of the callout line retain the fixed length returned by the **[Length](Excel.CalloutFormat.Length.md)** property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (**[MsoCalloutType](office.msocallouttype.md)** types **msoCalloutThree** and **msoCalloutFour**).


## Syntax

_expression_.**AutomaticLength**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

Applying this method sets the **[AutoLength](Excel.CalloutFormat.AutoLength.md)** property to **True**.


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