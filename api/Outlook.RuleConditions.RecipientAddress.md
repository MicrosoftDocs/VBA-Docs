---
title: RuleConditions.RecipientAddress property (Outlook)
keywords: vbaol11.chm2317
f1_keywords:
- vbaol11.chm2317
ms.prod: outlook
api_name:
- Outlook.RuleConditions.RecipientAddress
ms.assetid: 1b8f361e-0481-75dc-e66e-2bc69228773a
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.RecipientAddress property (Outlook)

Returns an **[AddressRuleCondition](Outlook.AddressRuleCondition.md)** object with an **[AddressRuleCondition.ConditionType](Outlook.AddressRuleCondition.ConditionType.md)** of **olConditionRecipientAddress**. Read-only.


## Syntax

_expression_. `RecipientAddress`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **AddressRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the recipient address contain the specified text.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **AddressRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[AddressRuleCondition.Enabled](Outlook.AddressRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]