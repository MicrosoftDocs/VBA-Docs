---
title: Chart.DataTable property (PowerPoint)
keywords: vbapp10.chm684003
f1_keywords:
- vbapp10.chm684003
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.DataTable
ms.assetid: fd212746-be95-06dd-144e-e6a4edf28e94
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DataTable property (PowerPoint)

Returns the chart data table. Read-only  **[DataTable](PowerPoint.DataTable.md)**.


## Syntax

_expression_.**DataTable**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example adds a data table with an outline border to the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.HasDataTable = True

        .Chart.DataTable.HasBorderOutline = True

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]