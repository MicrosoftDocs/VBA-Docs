---
title: OMathFunctions.Add method (Word)
keywords: vbawd10.chm44302440
f1_keywords:
- vbawd10.chm44302440
ms.prod: word
api_name:
- Word.OMathFunctions.Add
ms.assetid: 2292e297-6d24-cd73-971b-146be1edcb0a
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathFunctions.Add method (Word)

Inserts a new structure, such as a fraction, into an equation at the specified position and returns an **OMathFunction** object that represents the structure.


## Syntax

_expression_.**Add** (_Range_, _Type_, _NumArgs_, _NumCols_)

 _expression_ An expression that returns a [OMathFunctions](./Word.OMathFunctions.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**| The place at which to insert an equation.|
| _Type_|Required| **WdOMathFunctionType**|The type of equation to insert.|
| _NumArgs_|Optional| **Variant**| The number of arguments in the equation.|
| _NumCols_|Optional| **Variant**|The number of columns in the equation.|

## Return value

OMathFunction


## See also


[OMathFunctions Collection](Word.OMathFunctions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]