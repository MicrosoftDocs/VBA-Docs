---
title: CalloutFormat.Length property (Word)
keywords: vbawd10.chm163905644
f1_keywords:
- vbawd10.chm163905644
ms.prod: word
api_name:
- Word.CalloutFormat.Length
ms.assetid: 60b80a93-7a31-c4f6-57ab-445d788f6cbd
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat.Length property (Word)

Returns the length (in points) of the first segment of the callout line (the segment attached to the text callout box). Read-only  **Single**.


## Syntax

 _expression_. `Length`

 _expression_ An expression that returns a '[CalloutFormat](Word.CalloutFormat.md)' object.


## Remarks

The  **Length** property returns a value only when the **[AutoLength](Word.CalloutFormat.AutoLength.md)** property of the specified callout is set to **False** and applies only to callouts whose lines consist of more than one segment (types **msoCalloutThree** and **msoCalloutFour**).

This property is read-only. Use the  **[CustomLength](Word.CalloutFormat.CustomLength.md)** method to set the value of this property for the **[CalloutFormat](Word.CalloutFormat.md)** object.


## Example

This example specifies that if the first line segment in the callout named "co1" has a fixed length, then the length of the first line segment in the callout named "co2" will also be fixed at that same length. For the example to work, both callouts must have multiple-segment lines.


```vb
Dim sngLength As Single 
 
With ActiveDocument.Shapes 
 With .Item("co1").Callout 
 If Not .AutoLength Then sngLength = .Length 
 End With 
 If sngLength Then _ 
 .Item("co2").Callout.CustomLength sngLength 
End With
```


## See also


[CalloutFormat Object](Word.CalloutFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]