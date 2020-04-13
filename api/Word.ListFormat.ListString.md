---
title: ListFormat.ListString property (Word)
keywords: vbawd10.chm163577931
f1_keywords:
- vbawd10.chm163577931
ms.prod: word
api_name:
- Word.ListFormat.ListString
ms.assetid: b426ab7b-158a-0ae8-7c02-d71ef6a84263
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.ListString property (Word)

Returns a  **String** that represents the appearance of the list value of the first paragraph in the range for the specified **ListFormat** object. For example, the second paragraph in an alphabetical list would return B. Read-only.


## Syntax

_expression_. `ListString`

 _expression_ An expression that returns a '[ListFormat](Word.ListFormat.md)' object.


## Remarks

For a bulleted list, you will need to apply the correct font to see the string. Most bullets use the Symbol or Wingdings font.

Use the **[ListValue](Word.ListFormat.ListValue.md)** property to return the numeric value of the paragraph.


## Example

This example displays both the numeric value of the first paragraph in the selection and the string representation of the list value.


```vb
v = Selection.Range.ListFormat.ListValue 
lstring = Selection.Range.ListFormat.ListString 
MsgBox "List value " & v _ 
 & " is represented by the string " & lstring
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]