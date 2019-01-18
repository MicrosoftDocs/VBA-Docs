---
title: xlSparklineRowCol enumeration (Excel)
ms.prod: excel
api_name:
- Excel.XlSparklineRowCol
ms.assetid: 1b978b0d-c2a9-3367-cdef-429f79d84882
ms.date: 06/08/2017
localization_priority: Normal
---


# xlSparklineRowCol enumeration (Excel)

Specifies how to plot the sparkline when the data on which it is based is in a square-shaped range.



|Name|Value|Description|
|:-----|:-----|:-----|
| **SparklineColumnsSquare**|2|Plot the data by columns.|
| **SparklineNonSquare**|0|The sparkline is not bound to data in a square-shaped range.|
| **SparklineRowsSquare**|1|Plot the data by rows.|

## Remarks

The  **xlSparklineRowCol** enumeration is used by the **[PlotBy](./overview/Excel.md)** property of the **[SparklineGroup](Excel.SparklineGroup.md)** object to determine how to plot chart in a sparkline when data on which it based is in a square-shaped range, such as A1:B2.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]