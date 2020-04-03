---
title: Axis.MinorTickMark property (Word)
keywords: vbawd10.chm113049637
f1_keywords:
- vbawd10.chm113049637
ms.prod: word
api_name:
- Word.Axis.MinorTickMark
ms.assetid: 7e00472d-6e50-929b-c841-a36cd6c01782
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorTickMark property (Word)

Returns or sets the type of minor tick mark for the specified axis. Read/write  **[XlTickMark](Word.xltickmark.md)**.


## Syntax

_expression_. `MinorTickMark`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

 **MinorTickMark** can be one of the following **xlTickMark** constants:


-  **xlTickMarkInside**
    
-  **xlTickMarkOutside**
    
-  **xlTickMarkCross**
    
-  **xlTickMarkNone**
    

## Example

The following example sets the minor tick marks for the value axis of the first chart in the active document to be inside the axis.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).MinorTickMark = xlTickMarkInside 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]