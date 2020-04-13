---
title: RuleConditions.FormName property (Outlook)
keywords: vbaol11.chm2314
f1_keywords:
- vbaol11.chm2314
ms.prod: outlook
api_name:
- Outlook.RuleConditions.FormName
ms.assetid: 9f292443-1af7-500e-2959-1fce4c7d4824
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.FormName property (Outlook)

Returns a **[FormNameRuleCondition](Outlook.FormNameRuleCondition.md)** object with a **[FormNameRuleCondition.ConditionType](Outlook.FormNameRuleCondition.ConditionType.md)** of **olConditionFormName**. Read-only.


## Syntax

_expression_. `FormName`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **FormNameRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that the message uses a specified form.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **FormNameRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[FormNameRuleCondition.Enabled](Outlook.FormNameRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]