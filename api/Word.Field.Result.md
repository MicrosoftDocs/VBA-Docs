---
title: Field.Result property (Word)
keywords: vbawd10.chm154075140
f1_keywords:
- vbawd10.chm154075140
ms.prod: word
api_name:
- Word.Field.Result
ms.assetid: 97b754cf-6598-63d4-5314-c1bbfacc76ab
ms.date: 06/08/2017
localization_priority: Normal
---


# Field.Result property (Word)

Returns a  **Range** object that represents a field's result. Read/write.


## Syntax

_expression_. `Result`

_expression_ Required. A variable that represents a '[Field](Word.Field.md)' object.


## Remarks

You can access a field result without changing the view from field codes. Use the  **Text** property to return text from a **Range** object.


## Example

This example applies bold formatting to the first field in the selection.


```vb
If Selection.Fields.Count >= 1 Then 
 Set myRange = Selection.Fields(1).Result 
 myRange.Bold = True 
End If
```


## See also


[Field Object](Word.Field.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]