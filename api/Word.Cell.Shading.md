---
title: Cell.Shading property (Word)
keywords: vbawd10.chm156106857
f1_keywords:
- vbawd10.chm156106857
ms.prod: word
api_name:
- Word.Cell.Shading
ms.assetid: ab2f5789-ba6e-fa8a-d0a9-4c8b7922aa92
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Shading property (Word)

Returns a  **[Shading](Word.Shading.md)** object that refers to the shading formatting for the specified object.


## Syntax

_expression_. `Shading`

_expression_ A variable that represents a '[Cell](Word.Cell.md)' object.


## Example

This example applies horizontal line texture to the first cell in the first row in first table.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Cells(1).Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]