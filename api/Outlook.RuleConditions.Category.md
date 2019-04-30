---
title: RuleConditions.Category property (Outlook)
keywords: vbaol11.chm2313
f1_keywords:
- vbaol11.chm2313
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Category
ms.assetid: f1131bf8-4752-4e93-c68d-73c0511d22da
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.Category property (Outlook)

Returns a  **[CategoryRuleCondition](Outlook.CategoryRuleCondition.md)** object with a **[CategoryRuleCondition.ConditionType](Outlook.CategoryRuleCondition.ConditionType.md)** of **olConditionCategory**. Read-only.


## Syntax

_expression_.**Category**

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **CategoryRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message is assigned specific categories.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **CategoryRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[CategoryRuleCondition.Enabled](Outlook.CategoryRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]