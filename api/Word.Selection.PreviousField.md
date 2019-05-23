---
title: Selection.PreviousField method (Word)
keywords: vbawd10.chm158662833
f1_keywords:
- vbawd10.chm158662833
ms.prod: word
api_name:
- Word.Selection.PreviousField
ms.assetid: 9361a318-9ee2-fd72-9d52-106abfd8d44e
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.PreviousField method (Word)

Selects and returns the previous field.


## Syntax

_expression_. `PreviousField`

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Return value

Field


## Remarks

 If this method finds a field, it returns a **Field** object; if not, it returns **Nothing**.


## Example

This example updates the previous field (the field immediately preceding the selection).


```vb
If Not (Selection.PreviousField Is Nothing) Then 
 Selection.Fields.Update 
End If
```

This example selects the previous field, and if a field is found, displays a message in the status bar.




```vb
Set myField = Selection.PreviousField 
If Not (myField Is Nothing) Then StatusBar = "Field found"
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]