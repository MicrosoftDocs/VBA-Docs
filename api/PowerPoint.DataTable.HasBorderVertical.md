---
title: DataTable.HasBorderVertical property (PowerPoint)
keywords: vbapp10.chm698003
f1_keywords:
- vbapp10.chm698003
ms.prod: powerpoint
api_name:
- PowerPoint.DataTable.HasBorderVertical
ms.assetid: 943d7af7-e1fe-e7fe-408b-154fa2ad3705
ms.date: 06/08/2017
localization_priority: Normal
---


# DataTable.HasBorderVertical property (PowerPoint)

 **True** if the chart data table has vertical cell borders. Read/write **Boolean**.


## Syntax

_expression_.**HasBorderVertical**

_expression_ A variable that represents a '[DataTable](PowerPoint.DataTable.md)' object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the data table for the first chart in the active document to be displayed with an outline border and no cell borders.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart

            .HasDataTable = True

            With .DataTable

                .HasBorderHorizontal = False

                .HasBorderVertical = False

                .HasBorderOutline = True

            End With

        End With

    End If

End With
```


## See also


[DataTable Object](PowerPoint.DataTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]