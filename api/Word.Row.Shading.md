---
title: Row.Shading property (Word)
keywords: vbawd10.chm156237927
f1_keywords:
- vbawd10.chm156237927
ms.prod: word
api_name:
- Word.Row.Shading
ms.assetid: 79aee52a-8f9c-d41c-7247-2f7432f49683
ms.date: 06/08/2017
localization_priority: Normal
---


# Row.Shading property (Word)

Returns a  **[Shading](Word.Shading.md)** object that refers to the shading formatting for the specified object.


## Syntax

 _expression_. `Shading`

 _expression_ Required. A variable that represents a '[Row](Word.Row.md)' object.


## Example

This example applies horizontal line texture to the first row in table one.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 With ActiveDocument.Tables(1).Rows(1).Shading 
 .Texture = wdTextureHorizontal 
 End With 
End If
```


## See also


[Row Object](Word.Row.md)

