---
title: DataTable object (PowerPoint)
keywords: vbapp10.chm698000
f1_keywords:
- vbapp10.chm698000
ms.prod: powerpoint
api_name:
- PowerPoint.DataTable
ms.assetid: eaa7cdda-e374-7d19-47a6-87e4458fc244
ms.date: 06/08/2017
localization_priority: Normal
---


# DataTable object (PowerPoint)

Represents a chart data table.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[DataTable](PowerPoint.Chart.DataTable.md)** property to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.HasDataTable = True

        .Chart.DataTable.HasBorderOutline = True

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]