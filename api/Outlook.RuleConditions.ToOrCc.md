---
title: RuleConditions.ToOrCc property (Outlook)
keywords: vbaol11.chm2309
f1_keywords:
- vbaol11.chm2309
ms.prod: outlook
api_name:
- Outlook.RuleConditions.ToOrCc
ms.assetid: 28a34223-e47d-3843-4648-fe757568d406
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.ToOrCc property (Outlook)

Returns a  **[RuleCondition](Outlook.RuleCondition.md)** object with a **[RuleCondition.ConditionType](Outlook.RuleCondition.ConditionType.md)** of **olConditionToOrCc**. Read-only.


## Syntax

_expression_. `ToOrCc`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that your name is in the **To** or **Cc** box.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **RuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](Outlook.RuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]