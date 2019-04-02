---
title: Rule.ExecutionOrder property (Outlook)
keywords: vbaol11.chm2169
f1_keywords:
- vbaol11.chm2169
ms.prod: outlook
api_name:
- Outlook.Rule.ExecutionOrder
ms.assetid: 070d50ca-4b0b-5629-1609-81ab8a3620d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule.ExecutionOrder property (Outlook)

Returns or sets a  **Long** that indicates the order of execution of the rule among other rules in the **[Rules](Outlook.Rules.md)** collection. Read/write.


## Syntax

_expression_. `ExecutionOrder`

_expression_ A variable that represents a [Rule](Outlook.Rule.md) object.


## Remarks

 **ExecutionOrder** is directly mapped with the numerical value of _Index_ in the **[Item](Outlook.Rules.Item.md)** method. For example, `Rules.Item(1)` represents a rule with **ExecutionOrder** being 1, `Rules.Item(2)` represents a rule with **ExecutionOrder** being 2, and `Rules.Item(Rules.Count)` represents the rule with **ExecutionOrder** being **[Count](Outlook.Rules.Count.md)** property.


## See also


[Rule Object](Outlook.Rule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]