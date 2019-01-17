---
title: Walls object (Excel)
keywords: vbaxl10.chm613072
f1_keywords:
- vbaxl10.chm613072
ms.prod: excel
api_name:
- Excel.Walls
ms.assetid: 9c6f0c5b-dbb8-7d71-44b7-29987e750cd3
ms.date: 06/08/2017
localization_priority: Normal
---


# Walls object (Excel)

Represents the walls of a 3-D chart. This object isn't a collection. There's no object that represents a single wall; you must return all the walls as a unit.


## Example

Use the  **[Walls](Excel.Chart.Walls.md)** property to return the **Walls** object. The following example sets the pattern on the walls for embedded chart one on Sheet1. If the chart isn't a 3-D chart, this example will fail.


```vb
Worksheets("Sheet1").ChartObjects(1).Chart _ 
 .Walls.Interior.Pattern = xlGray75
```


## See also



[Excel Object Model Reference](./overview/Excel/object-model.md)

