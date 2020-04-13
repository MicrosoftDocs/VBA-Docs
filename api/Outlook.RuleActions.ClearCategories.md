---
title: RuleActions.ClearCategories property (Outlook)
keywords: vbaol11.chm3233
f1_keywords:
- vbaol11.chm3233
ms.prod: outlook
api_name:
- Outlook.RuleActions.ClearCategories
ms.assetid: db594b52-1700-67a7-8445-338f7df254e9
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.ClearCategories property (Outlook)

Returns a **[RuleAction](Outlook.RuleAction.md)** object with a **[RuleAction.ActionType](Outlook.RuleAction.ActionType.md)** of **olRuleActionClearCategories**. Read-only.


## Syntax

_expression_. `ClearCategories`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a rule that specifies removing all the categories of a message as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **RuleAction** object, regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](Outlook.RuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]