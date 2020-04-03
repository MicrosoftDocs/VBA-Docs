---
title: Axis.MajorTickMark property (Word)
keywords: vbawd10.chm113049618
f1_keywords:
- vbawd10.chm113049618
ms.prod: word
api_name:
- Word.Axis.MajorTickMark
ms.assetid: f2e4c509-0736-44bd-249b-1963ac697ee4
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MajorTickMark property (Word)

Returns or sets the type of major tick mark for the specified axis. Read/write  **[XlTickMark](Word.xltickmark.md)**.


## Syntax

_expression_. `MajorTickMark`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

 **MajorTickMark** can be set to one of the following **xlTickMark** constants:


-  **xlTickMarkInside**
    
-  **xlTickMarkOutside**
    
-  **xlTickMarkCross**
    
-  **xlTickMarkNone**
    

## Example

The following example sets the major tick marks for the value axis for the first chart in the active document to be outside the axis.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).MajorTickMark = xlTickMarkOutside 
 End If 
End With 

```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]