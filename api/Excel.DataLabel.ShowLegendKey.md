---
title: DataLabel.ShowLegendKey property (Excel)
keywords: vbaxl10.chm582096
f1_keywords:
- vbaxl10.chm582096
ms.prod: excel
api_name:
- Excel.DataLabel.ShowLegendKey
ms.assetid: 0857f78c-1c96-1887-e55e-4997dc22afb0
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowLegendKey property (Excel)

 **True** if the data label legend key is visible. Read/write **Boolean**.


## Syntax

_expression_. `ShowLegendKey`

_expression_ A variable that represents a [DataLabel](Excel.DataLabel-graph-property.md) object.


## Example

This example sets the data labels for series one in Chart1 to show values and the legend key.


```vb
With Charts("Chart1").SeriesCollection(1).DataLabels 
 .ShowLegendKey = True 
 .Type = xlShowValue 
End With
```


## See also


[DataLabel Object](Excel.DataLabel(object).md)

