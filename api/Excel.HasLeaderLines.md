---
title: HasLeaderLines property (Excel Graph)
keywords: vbagr10.chm66930
f1_keywords:
- vbagr10.chm66930
ms.prod: excel
api_name:
- Excel.HasLeaderLines
ms.assetid: 7bd99363-8845-f74f-2732-7570427d7a22
ms.date: 04/11/2019
localization_priority: Normal
---


# HasLeaderLines property (Excel Graph)

**True** if the series has leader lines. Read/write **Boolean**.

## Syntax

_expression_.**HasLeaderLines**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example adds data labels and blue leader lines to series one on the pie chart.

```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]