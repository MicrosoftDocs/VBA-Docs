---
title: SparklineGroup.ModifySourceData method (Excel)
keywords: vbaxl10.chm871080
f1_keywords:
- vbaxl10.chm871080
ms.prod: excel
api_name:
- Excel.SparklineGroup.ModifySourceData
ms.assetid: 35c1c1ed-b61d-2412-961f-8eb74b5563a2
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroup.ModifySourceData method (Excel)

Sets the range that represents the source data for the sparkline group.


## Syntax

_expression_.**ModifySourceData** (_SourceData_)

_expression_ A variable that represents a **[SparklineGroup](Excel.SparklineGroup.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SourceData_|Required| **String**|The range that represents the source data.|

## Return value

Nothing


## Example

This example selects a sparkline group in the location A1:A4, and modifies the source data to include an additional column by using the data in the range B1:D4.

```vb
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifySourceData "B1:D4"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]