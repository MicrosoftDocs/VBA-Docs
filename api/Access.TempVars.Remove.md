---
title: TempVars.Remove method (Access)
keywords: vbaac10.chm14070
f1_keywords:
- vbaac10.chm14070
ms.prod: access
api_name:
- Access.TempVars.Remove
ms.assetid: a9ab9ff2-5bfc-d001-f5eb-9929907bc1b2
ms.date: 03/26/2019
localization_priority: Normal
---


# TempVars.Remove method (Access)

Removes the specified **[TempVar](Access.TempVar.md)** object from the **TempVars** collection.


## Syntax

_expression_.**Remove** (_var_)

_expression_ A variable that represents a **[TempVars](Access.TempVars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _var_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the _expression_ argument. If a numeric expression, the argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, the argument must be the name of a member of the collection.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]