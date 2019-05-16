---
title: SparklineGroup.Location property (Excel)
keywords: vbaxl10.chm871076
f1_keywords:
- vbaxl10.chm871076
ms.prod: excel
api_name:
- Excel.SparklineGroup.Location
ms.assetid: 3548cc42-dbab-636f-0dcf-2f38ad4a2db5
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroup.Location property (Excel)

Gets or sets the **[Range](Excel.Range(object).md)** object that represents the location of the sparkline group. Read/write.


## Syntax

_expression_.**Location** 

_expression_ A variable that represents a **[SparklineGroup](Excel.SparklineGroup.md)** object.


## Return value

**Range**


## Remarks

The location for all sparklines in a sparkline group must be on the same sheet, but the source data for the sparkline group can be on a different sheet or workbook.

The size of the range that represents the **Location** property must equal the number of rows or columns in the **[SourceData](Excel.SparklineGroup.SourceData.md)** property.

A continuous range associated with a sparkline group can only be one-dimensional. If the range is not continuous, each cell must be specified individually.

> [!NOTE] 
> Do not use the **[Union](Excel.Application.Union.md)** method to create a non-continuous range because the **Union** method returns a single range reference.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]