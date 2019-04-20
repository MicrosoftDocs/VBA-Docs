---
title: Chart.SaveChartTemplate method (Excel)
keywords: vbaxl10.chm149181
f1_keywords:
- vbaxl10.chm149181
ms.prod: excel
api_name:
- Excel.Chart.SaveChartTemplate
ms.assetid: d9e36023-b5bb-aaf4-5b34-9a22df468ced
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.SaveChartTemplate method (Excel)

Saves a custom chart template to the list of available chart templates.


## Syntax

_expression_.**SaveChartTemplate** (_FileName_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the chart template.|

## Remarks

By default, this method saves the active chart to the user's chart template directory. If a UNC or URL is specified, the chart will be saved to the specified location instead. 


## Example

This example adds a new chart template based on the active chart.

```vb
ActiveChart.SaveChartTemplate _ 
 Filename:="Presentation Chart" 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]