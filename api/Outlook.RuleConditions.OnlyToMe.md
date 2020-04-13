---
title: RuleConditions.OnlyToMe property (Outlook)
keywords: vbaol11.chm2307
f1_keywords:
- vbaol11.chm2307
ms.prod: outlook
api_name:
- Outlook.RuleConditions.OnlyToMe
ms.assetid: 208e7bc4-2938-ecc8-7af5-9e3e256fe5b1
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.OnlyToMe property (Outlook)

Returns a **[RuleCondition](Outlook.RuleCondition.md)** object with a **[RuleCondition.ConditionType](Outlook.RuleCondition.ConditionType.md)** of **olConditionOnlyToMe**. Read-only.


## Syntax

_expression_. `OnlyToMe`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message is sent only to you.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](Outlook.RuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]