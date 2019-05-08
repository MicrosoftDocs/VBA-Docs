---
title: ChartArea.ClearFormats method (Word)
keywords: vbawd10.chm160039024
f1_keywords:
- vbawd10.chm160039024
ms.prod: word
api_name:
- Word.ChartArea.ClearFormats
ms.assetid: 4a528ed5-dec3-13f9-3a83-b3dcdfe79329
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea.ClearFormats method (Word)

Clears the formatting of the object.


## Syntax

_expression_.**ClearFormats**

_expression_ A variable that represents a '[ChartArea](Word.ChartArea.md)' object.


## Example

The following example clears the formatting from the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartArea.ClearFormats 
 End If 
End With
```


## See also


[ChartArea Object](Word.ChartArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]