---
title: RuleConditions.FromAnyRSSFeed property (Outlook)
keywords: vbaol11.chm3250
f1_keywords:
- vbaol11.chm3250
ms.prod: outlook
api_name:
- Outlook.RuleConditions.FromAnyRSSFeed
ms.assetid: df580ca7-ee2f-9c3a-ebc7-ca35528554cd
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.FromAnyRSSFeed property (Outlook)

Returns a **[RuleCondition](Outlook.RuleCondition.md)** object with a **[RuleCondition.ConditionType](Outlook.RuleCondition.ConditionType.md)** of **olConditionFromAnyRssFeed**. Read-only.


## Syntax

_expression_. `FromAnyRSSFeed`

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

Use the returned  **RuleCondition** object when enumerating the rule conditions or exception conditions of an existing rule, or when creating a rule that specifies the condition or exception condition that the message is from an RSS subscription.

This property of the  **[RuleConditions](Outlook.RuleConditions.md)** collection always returns a **RuleCondition** object, regardless of whether the rule associated with this **RuleConditions** collection has defined such a rule condition. If the rule has defined and enabled such a rule condition, then **[RuleCondition.Enabled](Outlook.RuleCondition.Enabled.md)** will be **True**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]