---
title: ReturnVars.Item property (Access)
keywords: vbaac10.chm14695
f1_keywords:
- vbaac10.chm14695
ms.prod: access
api_name:
- Access.ReturnVars.Item
ms.assetid: b6bdac31-e39d-434b-8ec4-28cee04389b8
ms.date: 03/06/2019
localization_priority: Normal
---


# ReturnVars.Item property (Access)

The **Item** property returns a specific member of a collection either by position or by index. Read-only **ReturnVar**.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[ReturnVars](Access.ReturnVars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|An expression that specifies the position of a member of the collection referred to by the _expression_ argument.<br/><br/>If a numeric expression, the _Index_ argument must be a number from 0 to the value of the collection's **Count** property minus 1.<br/><br/>If a string expression, the _Index_ argument must be the name of a member of the collection.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]