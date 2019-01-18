---
title: Chart.SetDefaultChart method (Excel)
keywords: vbaxl10.chm149182
f1_keywords:
- vbaxl10.chm149182
ms.prod: excel
api_name:
- Excel.Chart.SetDefaultChart
ms.assetid: 8be43de3-8b7d-4885-3e49-19aa0c65564f
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SetDefaultChart method (Excel)

Specifies the name of the chart template that Microsoft Excel uses when creating new charts.


## Syntax

_expression_. `SetDefaultChart`( `_Name_` )

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **Variant**|Specifies the name of the default chart template that will be used when creating new charts. This name can be a string naming a chart in the gallery for a user-defined template or it can be a special constant  **xlBuiltIn** to specify a built-in chart template.|

## Example

This example sets the default chart template to the custom chart named "Monthly Sales."


```vb
ActiveChart.SetDefaultChart Name:="Monthly Sales"
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]