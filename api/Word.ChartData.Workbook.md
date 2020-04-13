---
title: ChartData.Workbook property (Word)
keywords: vbawd10.chm190382080
f1_keywords:
- vbawd10.chm190382080
ms.prod: word
api_name:
- Word.ChartData.Workbook
ms.assetid: 2295d653-7a36-b258-dfb8-f48844331705
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.Workbook property (Word)

Returns the workbook that contains the chart data associated with the chart. Read-only  **Object**.


## Syntax

_expression_.**Workbook**

_expression_ A variable that represents a '[ChartData](Word.ChartData.md)' object.


## Remarks




> [!NOTE] 
> You must call the **[Activate](Word.ChartData.Activate.md)** method before referencing this property; otherwise, an error occurs.


## Example

The following example activates the Microsoft Excel workbook associated with the first chart in the active document. If the Excel workbook has multiple windows, the example activates the first window. The example then copies the contents of cells B1 through B5 and pastes the cell contents into the chart.


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


[ChartData Object](Word.ChartData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]