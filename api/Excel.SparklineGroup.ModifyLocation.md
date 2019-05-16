---
title: SparklineGroup.ModifyLocation method (Excel)
keywords: vbaxl10.chm871079
f1_keywords:
- vbaxl10.chm871079
ms.prod: excel
api_name:
- Excel.SparklineGroup.ModifyLocation
ms.assetid: 8f6ca2cb-b0cc-a0bf-efc0-ee30ca3888e6
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroup.ModifyLocation method (Excel)

Sets the associated **[Range](excel.range(object).md)** object to modify the location of the sparkline group.

## Syntax

_expression_.**ModifyLocation** (_Location_)

_expression_ A variable that represents a **[SparklineGroup](Excel.SparklineGroup.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **Range**|The range that represents the location of the sparkline group.|

## Return value

Nothing

## Example

This example selects a sparkline group in the location A1:A4 and changes the location to equal A10:A14.

```vb
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifyLocation Range("$A$10:$A$14")
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]