---
title: HasDataLabels property (Excel Graph)
keywords: vbagr10.chm65614
f1_keywords:
- vbagr10.chm65614
ms.prod: excel
api_name:
- Excel.HasDataLabels
ms.assetid: 1aa1d13e-69ec-0dab-1820-437c09afe820
ms.date: 04/11/2019
localization_priority: Normal
---


# HasDataLabels property (Excel Graph)

**True** if the series has data labels. Read/write **Boolean**.

## Syntax

_expression_.**HasDataLabels**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example turns on data labels for series three.

```vb
With myChart.SeriesCollection(3) 
 .HasDataLabels = True 
 .ApplyDataLabels Type:=xlValue 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]