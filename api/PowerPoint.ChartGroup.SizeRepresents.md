---
title: ChartGroup.SizeRepresents property (PowerPoint)
keywords: vbapp10.chm692001
f1_keywords:
- vbapp10.chm692001
ms.prod: powerpoint
api_name:
- PowerPoint.ChartGroup.SizeRepresents
ms.assetid: 69570b42-c850-1381-d18d-d924bd30352a
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SizeRepresents property (PowerPoint)

Returns or sets what the bubble size represents on a bubble chart. Read/write  **Long**.


## Syntax

_expression_.**SizeRepresents**

_expression_ A variable that represents a **[ChartGroup](PowerPoint.ChartGroup.md)** object.


## Remarks

This property can be either of the following  **[XlSizeRepresents](PowerPoint.XlSizeRepresents.md)** constants:


-  **xlSizeIsArea**
    
-  **xlSizeIsWidth**
    



## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets what the bubble size represents for chart group one of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartGroups(1).SizeRepresents = xlSizeIsWidth

    End If

End With
```


## See also


[ChartGroup Object](PowerPoint.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]