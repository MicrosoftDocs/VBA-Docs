---
title: SparklineGroups.Group method (Excel)
keywords: vbaxl10.chm869080
f1_keywords:
- vbaxl10.chm869080
ms.prod: excel
api_name:
- Excel.SparklineGroups.Group
ms.assetid: a5e01669-1922-4b26-158d-3c3aa70a101a
ms.date: 05/16/2019
localization_priority: Normal
---


# SparklineGroups.Group method (Excel)

Groups the selected sparklines.


## Syntax

_expression_.**Group** (_Location_)

_expression_ A variable that represents a **[SparklineGroups](Excel.SparklineGroups.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Location_|Required| **[Range](Excel.Range(object).md)**|The location of the first cell in the group.|

## Return value

Nothing


## Example

This example selects the range A1:A4 and groups the sparklines in that range.

```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Group Location:=Range("A1")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]