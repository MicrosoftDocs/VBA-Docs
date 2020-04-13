---
title: RuleConditions.OnLocalMachine property (Outlook)
keywords: vbaol11.chm2322
f1_keywords:
- vbaol11.chm2322
ms.prod: outlook
api_name:
- Outlook.RuleConditions.OnLocalMachine
ms.assetid: 747de02c-d76d-9da3-c582-50719e618eb4
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.OnLocalMachine property (Outlook)

Returns a **[RuleCondition](Outlook.RuleCondition.md)** object with a **[RuleCondition.ConditionType](Outlook.RuleCondition.ConditionType.md)** of **olConditionLocalMachineOnly**. Read-only.


## Syntax

_expression_. `OnLocalMachine`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the rule can run on this machine only. When you run the same rule on another computer, the rule will show that the condition **olConditionOtherMachine** is enabled.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. You can programmatically enable a rule condition of this type. In certain cases, **olConditionLocalMachine** is automatically set as a result of enabling another rule condition such as **olConditionAccount**. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](Outlook.RuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]