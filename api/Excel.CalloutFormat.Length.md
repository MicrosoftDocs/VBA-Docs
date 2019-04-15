---
title: CalloutFormat.Length property (Excel)
keywords: vbaxl10.chm104014
f1_keywords:
- vbaxl10.chm104014
ms.prod: excel
api_name:
- Excel.CalloutFormat.Length
ms.assetid: e17dacaa-f48f-8802-3912-f84a0e4dd8ca
ms.date: 04/13/2019
localization_priority: Normal
---


# CalloutFormat.Length property (Excel)

Returns a **Single** value that represents the length (in [points](../language/glossary/vbe-glossary.md#point)) of the first segment of the callout line (the segment attached to the text callout box).


## Syntax

_expression_.**Length**

_expression_ A variable that represents a **[CalloutFormat](Excel.CalloutFormat.md)** object.


## Remarks

This property is read-only and applies only to callouts whose lines consist of more than one segment (**[MsoCalloutType](office.msocallouttype.md)** types **msoCalloutThree** and **msoCalloutFour**), and for which the **[AutoLength](Excel.CalloutFormat.AutoLength.md)** property is set to **False**. 

Use the **[CustomLength](Excel.CalloutFormat.CustomLength.md)** method to set the value of this property for a **CalloutFormat** object.


## Example

If the first line segment in the callout named callout1 has a fixed length, this example specifies that the length of the first line segment in the callout named callout2 on worksheet one will also be fixed at that length. For the example to work, both callouts must have multiple-segment lines.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
 With .Item("callout1").Callout 
 If Not .AutoLength Then len1 = .Length 
 End With 
 If len1 Then .Item("callout2").Callout.CustomLength len1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]