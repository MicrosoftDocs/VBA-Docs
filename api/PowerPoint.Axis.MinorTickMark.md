---
title: Axis.MinorTickMark property (PowerPoint)
keywords: vbapp10.chm682022
f1_keywords:
- vbapp10.chm682022
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.MinorTickMark
ms.assetid: 2486a649-7006-388f-1b52-379b44f3f80d
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorTickMark property (PowerPoint)

Returns or sets the type of minor tick mark for the specified axis. Read/write  **[XlTickMark](PowerPoint.XlTickMark.md)**.


## Syntax

_expression_. `MinorTickMark`

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Remarks

 **MinorTickMark** can be one of the following **xlTickMark** constants:


-  **xlTickMarkInside**
    
-  **xlTickMarkOutside**
    
-  **xlTickMarkCross**
    
-  **xlTickMarkNone**
    

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the minor tick marks for the value axis of the first chart in the active document to be inside the axis.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Axes(xlValue).MinorTickMark = xlTickMarkInside

    End If

End With
```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]