---
title: DataLabels method (Excel Graph)
keywords: vbagr10.chm3077616
f1_keywords:
- vbagr10.chm3077616
ms.prod: excel
api_name:
- Excel.DataLabels
ms.assetid: 8ffca32c-f505-482e-dd27-d29ad2682daf
ms.date: 04/09/2019
localization_priority: Normal
---


# DataLabels method (Excel Graph)

Returns an object that represents either a single data label (a **[DataLabel](Excel.DataLabel-graph-object.md)** object) or a collection of all the data labels for the series (a **[DataLabels](Excel.datalabels(collection).md)** collection).

## Syntax

_expression_.**DataLabels** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_| Optional |**Variant**|The number of the data label.|

## Example

This example sets the data labels for series one to show their key, assuming that their values are visible when the example runs.

```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 With .DataLabels 
 .ShowLegendKey = True 
 .Type = xlValue 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]