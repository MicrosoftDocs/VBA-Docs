---
title: TempVars.Add method (Access)
keywords: vbaac10.chm14069
f1_keywords:
- vbaac10.chm14069
ms.prod: access
api_name:
- Access.TempVars.Add
ms.assetid: 836e449c-35ff-4089-857a-403c9fc97592
ms.date: 06/08/2017
---


# TempVars.Add method (Access)

Adds a variable to the  **[TempVars](Access.TempVars.md)** collection.


## Syntax

_expression_. `Add`( ` _Name_`, ` _Value_` )

_expression_ A variable that represents a [TempVars](Access.TempVars.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name to use for the  **TempVar**.|
| _Value_|Required|**Variant**|The value to store as a  **TempVar**. This value must be a string expression or a numeric expression. Setting this argument to an object data type will result in a run-time error.|

## See also


[TempVars Collection](Access.TempVars.md)

