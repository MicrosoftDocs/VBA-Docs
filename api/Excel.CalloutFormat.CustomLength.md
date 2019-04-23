---
title: CalloutFormat.CustomLength method (Excel)
keywords: vbaxl10.chm104004
f1_keywords:
- vbaxl10.chm104004
ms.prod: excel
api_name:
- Excel.CalloutFormat.CustomLength
ms.assetid: 8c5034f9-32ca-6e34-be59-51e0cd8c8374
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.CustomLength method (Excel)

Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. 

Use the **[AutomaticLength](Excel.CalloutFormat.AutomaticLength.md)** method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (**[MsoCalloutType](office.msocallouttype.md)** types **msoCalloutThree** and **msoCalloutFour**).


## Syntax

_expression_.**CustomLength** (_Length_)

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Length_|Required| **Single**|The length of the first segment of the callout, in [points](../language/glossary/vbe-glossary.md#point).|

## Remarks

Applying this method sets the **[AutoLength](Excel.CalloutFormat.AutoLength.md)** property to **False** and sets the **[Length](Excel.CalloutFormat.Length.md)** property to the value specified for the _Length_ argument.


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