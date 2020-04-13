---
title: RuleConditions.Account property (Outlook)
keywords: vbaol11.chm2310
f1_keywords:
- vbaol11.chm2310
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Account
ms.assetid: 9e1ecf7d-b832-e657-92df-42bb28f5d924
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.Account property (Outlook)

Returns a **[AccountRuleCondition](Outlook.AccountRuleCondition.md)** object with an **[AccountRuleCondition.ConditionType](Outlook.AccountRuleCondition.ConditionType.md)** of **olConditionAccount**. Read-only.


## Syntax

_expression_. `Account`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **AccountRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that a message is sent or received through the specified account.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns an **AccountRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[AccountRuleCondition.Enabled](Outlook.AccountRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]