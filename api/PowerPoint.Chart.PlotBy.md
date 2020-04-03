---
title: Chart.PlotBy property (PowerPoint)
keywords: vbapp10.chm65738
f1_keywords:
- vbapp10.chm65738
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.PlotBy
ms.assetid: 14b696d7-148c-267f-4294-4dddc9fba4e1
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotBy property (PowerPoint)

Returns or sets the way columns or rows are used as data series on the chart. Read/write  **Long**.


## Syntax

_expression_.**PlotBy**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

The value of this property can be one of the following  **[XlRowCol](PowerPoint.XlRowCol.md)** constants:


-  **xlColumns**
    
-  **xlRows**
    


For PivotChart reports, this property is read-only and always returns  **xlColumns**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the first chart in the active document to plot data by columns.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.PlotBy = xlColumns

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]