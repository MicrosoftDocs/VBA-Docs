---
title: RuleConditions.From property (Outlook)
keywords: vbaol11.chm2315
f1_keywords:
- vbaol11.chm2315
ms.prod: outlook
api_name:
- Outlook.RuleConditions.From
ms.assetid: 3ebda0d0-ba44-95c6-ed02-a9c6acbf1f1c
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.From property (Outlook)

Returns a  **[ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md)** object with a **[ToOrFromRuleCondition.ConditionType](Outlook.ToOrFromRuleCondition.ConditionType.md)** of **olConditionFrom**. Read-only.


## Syntax

_expression_. `From`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **ToOrFromRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the sender of the message is in the specified list of **[Recipients](Outlook.Recipients.md)**.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **ToOrFromRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[ToOrFromRuleCondition.Enabled](Outlook.ToOrFromRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]