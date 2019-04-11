---
title: HasDataLabel property (Excel Graph)
keywords: vbagr10.chm5207462
f1_keywords:
- vbagr10.chm5207462
ms.prod: excel
api_name:
- Excel.HasDataLabel
ms.assetid: d8fd8c48-4723-4da9-0b8a-82d02c93a19f
ms.date: 04/11/2019
localization_priority: Normal
---


# HasDataLabel property (Excel Graph)

**True** if the point has a data label. Read/write **Boolean**.

## Syntax

_expression_.**HasDataLabel**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example turns on the data label for point seven in series three, and then it sets the data label color to blue.

```vb
With myChart.SeriesCollection(3).Points(7) 
 .HasDataLabel = True 
 .ApplyDataLabels Type:=xlValue 
 .DataLabel.Font.ColorIndex = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]