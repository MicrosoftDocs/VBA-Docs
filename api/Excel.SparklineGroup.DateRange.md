---
title: SparklineGroup.DateRange property (Excel)
keywords: vbaxl10.chm871078
f1_keywords:
- vbaxl10.chm871078
ms.prod: excel
api_name:
- Excel.SparklineGroup.DateRange
ms.assetid: 4944aa78-89cc-8252-2c5e-148ca4229579
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroup.DateRange property (Excel)

Gets or sets the date range for the sparkline group. Read/write.


## Syntax

_expression_.**DateRange**

_expression_ A variable that represents a **[SparklineGroup](Excel.SparklineGroup.md)** object.


## Return value

String


## Remarks

To clear the date range, set this property to an empty string.

The date range must be a continuous one-dimensional range.

The date range can be located on a different sheet than the **[Location](Excel.SparklineGroup.Location.md)** and **[SourceData](Excel.SparklineGroup.SourceData.md)** properties.

Empty cells and non-date values in the date range are not displayed.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]