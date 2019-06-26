---
title: Walls object (PowerPoint)
keywords: vbapp10.chm723000
f1_keywords:
- vbapp10.chm723000
ms.prod: powerpoint
api_name:
- PowerPoint.Walls
ms.assetid: b2288a5f-efec-84b4-9a40-d62d61196ac8
ms.date: 06/08/2017
localization_priority: Normal
---


# Walls object (PowerPoint)

Represents the walls of a 3D chart. 


## Remarks

This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Walls](PowerPoint.Chart.Walls.md)** property to return the **Walls** object. The following example sets the pattern on the walls for the first chart in the active document. If the chart is not a 3D chart, this example will fail.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Walls.Interior.Pattern = xlGray75

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]