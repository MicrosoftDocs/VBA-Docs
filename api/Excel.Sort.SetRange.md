---
title: Sort.SetRange method (Excel)
keywords: vbaxl10.chm847079
f1_keywords:
- vbaxl10.chm847079
ms.prod: excel
api_name:
- Excel.Sort.SetRange
ms.assetid: 12a68fb7-379d-f9fa-d464-a6d5fe1e6f9b
ms.date: 05/16/2019
localization_priority: Normal
---


# Sort.SetRange method (Excel)

Sets the range over which the sort occurs.


## Syntax

_expression_.**SetRange** (_Rng_)

_expression_ A variable that represents a **[Sort](Excel.Sort.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Rng_|Required| **Range**|Specifies the range over which the sort represented by the **Sort** object occurs.|


> [!NOTE] 
> **SetRange** can only be used when applying a sort to a sheet range, and cannot be used if the range is within a table.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]