---
title: RuleConditions.BodyOrSubject property (Outlook)
keywords: vbaol11.chm2312
f1_keywords:
- vbaol11.chm2312
ms.prod: outlook
api_name:
- Outlook.RuleConditions.BodyOrSubject
ms.assetid: ced8a26a-9a54-d1f4-18f6-dd52a8203892
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.BodyOrSubject property (Outlook)

Returns a **[TextRuleCondition](Outlook.TextRuleCondition.md)** object with a **[TextRuleCondition.ConditionType](Outlook.TextRuleCondition.ConditionType.md)** of **olConditionBodyOrSubject**. Read-only.


## Syntax

_expression_. `BodyOrSubject`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **TextRuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a new rule that specifies the condition or exception condition that certain text is in the message body or subject.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **TextRuleCondition** object regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[TextRuleCondition.Enabled](Outlook.TextRuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]