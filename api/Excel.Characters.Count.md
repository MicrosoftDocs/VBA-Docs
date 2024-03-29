---
title: Characters.Count property (Excel)
keywords: vbaxl10.chm252074
f1_keywords:
- vbaxl10.chm252074
api_name:
- Excel.Characters.Count
ms.assetid: 0fabbbe3-5c4a-c215-1bc0-201ee5971fb0
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Characters.Count property (Excel)

Returns a **Long** value that represents the number of objects in the collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[Characters](Excel.Characters.md)** object.


## Example

This example makes the last character in cell A1 a superscript character.

```vb
Sub MakeSuperscript() 
 Dim n As Integer 
 
 n = Worksheets("Sheet1").Range("A1").Characters.Count 
 Worksheets("Sheet1").Range("A1").Characters(n, 1) _ 
 .Font.Superscript = True 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
