---
title: Table.Shading property (Word)
keywords: vbawd10.chm156303464
f1_keywords:
- vbawd10.chm156303464
ms.prod: word
api_name:
- Word.Table.Shading
ms.assetid: 0c5c0ebe-d7cb-ff55-c77c-2c0c36a6c98a
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Shading property (Word)

Returns a  **Shading** object that refers to the shading formatting for the specified object.


## Syntax

_expression_. `Shading`

_expression_ Required. A variable that represents a '[Table](Word.Table.md)' object.


## Example

This example applies horizontal line texture to the first table in the active document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1)Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]