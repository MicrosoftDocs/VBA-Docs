---
title: RuleActions.NewItemAlert property (Outlook)
keywords: vbaol11.chm2199
f1_keywords:
- vbaol11.chm2199
ms.prod: outlook
api_name:
- Outlook.RuleActions.NewItemAlert
ms.assetid: 01de8523-7617-c3df-39c6-395f85eda57f
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.NewItemAlert property (Outlook)

Returns a **[NewItemAlertRuleAction](Outlook.NewItemAlertRuleAction.md)** object with **[ActionType](Outlook.NewItemAlertRuleAction.ActionType.md)** being **olRuleActionNewItemAlert**. Read-only.


## Syntax

_expression_. `NewItemAlert`

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **NewItemAlertRuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies displaying an alert for a new item as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **NewItemAlertRuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[NewItemAlertRuleAction.Enabled](Outlook.NewItemAlertRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]