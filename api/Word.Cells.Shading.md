---
title: Cells.Shading property (Word)
keywords: vbawd10.chm155844709
f1_keywords:
- vbawd10.chm155844709
ms.prod: word
api_name:
- Word.Cells.Shading
ms.assetid: ea9f4c8a-254d-6197-0f90-fa79465f940f
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.Shading property (Word)

Returns a  **[Shading](Word.Shading.md)** object that refers to the shading formatting for the specified object.


## Syntax

_expression_. `Shading`

_expression_ A variable that represents a '[Cells](Word.cells.md)' object.


## Example

This example applies horizontal line texture to the first row in table one.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Cells.Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]