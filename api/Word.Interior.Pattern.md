---
title: Interior.Pattern property (Word)
keywords: vbawd10.chm2818054
f1_keywords:
- vbawd10.chm2818054
ms.prod: word
api_name:
- Word.Interior.Pattern
ms.assetid: 5910e6a3-9aaa-7908-aa7d-345bdbabc4de
ms.date: 06/08/2017
localization_priority: Normal
---


# Interior.Pattern property (Word)

Returns or sets a  **Variant** value, containing an **[XlPattern](Word.xlpattern.md)** constant, that represents the interior pattern.


## Syntax

_expression_.**Pattern**

_expression_ A variable that represents an '[Interior](Word.Interior.md)' object.


## Example

The following example enables up and down bars, then adds a criss-cross pattern to the down bars and sets the pattern color to red, for the first chart group of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.Pattern = xlPatternCrissCross 
 .DownBars.Interior.PatternColorIndex = 3 
 End With 
 End If 
End With
```


## See also


[Interior Object](Word.Interior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]