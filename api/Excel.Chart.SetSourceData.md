---
title: Chart.SetSourceData method (Excel)
keywords: vbaxl10.chm149162
f1_keywords:
- vbaxl10.chm149162
ms.prod: excel
api_name:
- Excel.Chart.SetSourceData
ms.assetid: fc41cc05-087a-f53c-2f54-fd6307de51d6
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.SetSourceData method (Excel)

Sets the source data range for the chart.


## Syntax

_expression_.**SetSourceData** (_Source_, _PlotBy_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Range**|The range that contains the source data.|
| _PlotBy_|Optional| **Variant**|Specifies the way the data is to be plotted. Can be either of the following **[XlRowCol](Excel.XlRowCol.md)** constants: **xlColumns** or **xlRows**.|

## Example

This example sets the source data range for chart one.

```vb
Charts(1).SetSourceData Source:=Sheets(1).Range("a1:a10"), _ 
 PlotBy:=xlColumns
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
