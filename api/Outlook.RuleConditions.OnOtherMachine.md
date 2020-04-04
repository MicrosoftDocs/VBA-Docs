---
title: RuleConditions.OnOtherMachine property (Outlook)
keywords: vbaol11.chm2323
f1_keywords:
- vbaol11.chm2323
ms.prod: outlook
api_name:
- Outlook.RuleConditions.OnOtherMachine
ms.assetid: 03d96697-5978-8a0c-7356-dfe721f5b05d
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.OnOtherMachine property (Outlook)

Returns a **[RuleCondition](Outlook.RuleCondition.md)** object with a **[RuleCondition.ConditionType](Outlook.RuleCondition.ConditionType.md)** of **olConditionOtherMachine**. Read-only.


## Syntax

_expression_. `OnOtherMachine`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule. This rule condition indicates that the rule can run only on some machine other than the local machine.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition.

You cannot programmatically enable or disable a condition of type  **olConditionOtherMachine**. This type of rule condition indicates that the rule can run only on a specific computer that is not the current one. This happens when the rule is created on that computer and the rule condition **olConditionLocalMachineOnly** is enabled, indicating that the rule can run only on that computer. When you run the same rule on another computer, the rule will show that the condition **olConditionOtherMachine** is enabled.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]