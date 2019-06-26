---
title: ChartData.Workbook property (PowerPoint)
keywords: vbapp10.chm689001
f1_keywords:
- vbapp10.chm689001
ms.prod: powerpoint
api_name:
- PowerPoint.ChartData.Workbook
ms.assetid: 2d22aa4a-15d8-c5f3-5059-a968e9a85789
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.Workbook property (PowerPoint)

Returns the workbook that contains the chart data associated with the chart. Read-only **Object**.


## Syntax

_expression_.**Workbook**

_expression_ A variable that represents a '[ChartData](PowerPoint.ChartData.md)' object.


## Remarks




> [!NOTE] 
> You must call the **[Activate](PowerPoint.ChartData.Activate.md)** method before referencing this property; otherwise, an error occurs.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example activates the Microsoft Excel workbook associated with the first chart in the active document. If the Excel workbook has multiple windows, the example activates the first window. The example then copies the contents of cells B1 through B5 and pastes the cell contents into the chart.


> [!NOTE] 
> Excel must be open to modify data in the workbook.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then
        .Chart.ChartData.Activate
        .Chart.ChartData.Workbook. _
            Worksheets("Sheet1").Range("B1:B5").Copy
        .Chart.Paste
    End If

End With
```


## See also


[ChartData Object](PowerPoint.ChartData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]