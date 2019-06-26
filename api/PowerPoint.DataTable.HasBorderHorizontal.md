---
title: DataTable.HasBorderHorizontal property (PowerPoint)
keywords: vbapp10.chm698002
f1_keywords:
- vbapp10.chm698002
ms.prod: powerpoint
api_name:
- PowerPoint.DataTable.HasBorderHorizontal
ms.assetid: 6fb381e0-f990-656d-89e7-cb2f43a84ece
ms.date: 06/08/2017
localization_priority: Normal
---


# DataTable.HasBorderHorizontal property (PowerPoint)

 **True** if the chart data table has horizontal cell borders. Read/write **Boolean**.


## Syntax

_expression_.**HasBorderHorizontal**

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