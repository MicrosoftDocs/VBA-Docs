---
title: ChartData.Activate method (Word)
keywords: vbawd10.chm190382081
f1_keywords:
- vbawd10.chm190382081
ms.prod: word
api_name:
- Word.ChartData.Activate
ms.assetid: 08f4a657-41c2-52ea-b31c-976549ace8c1
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.Activate method (Word)

Activates the first window of the workbook associated with the chart.


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a '[ChartData](Word.ChartData.md)' object.


## Remarks

If the chart is linked to a Microsoft Excel workbook, this method does not run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run those macros).


> [!NOTE] 
> You must call this method before referencing the **[Workbook](Word.ChartData.Workbook.md)** property.


## Example

The following example activates the Excel workbook associated with the first chart in the active document. If the Excel workbook has multiple windows, the example activates the first window. The example then copies the contents of cells B1 through B5 and pastes the cell contents into the chart.


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