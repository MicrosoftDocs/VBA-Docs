---
title: RuleActions.Delete property (Outlook)
keywords: vbaol11.chm2186
f1_keywords:
- vbaol11.chm2186
ms.prod: outlook
api_name:
- Outlook.RuleActions.Delete
ms.assetid: eb219d46-64c2-650c-ad39-e049ef33160f
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.Delete property (Outlook)

Returns a  **[RuleAction](Outlook.RuleAction.md)** object with **[RuleAction.ActionType](Outlook.RuleAction.ActionType.md)** being **olRuleActionDelete**. Read-only.


## Syntax

_expression_.**Delete**

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

Use the returned  **RuleAction** object when enumerating the rule actions of an existing rule or when creating a new rule that specifies deleting a message as an action.

This property of the  **[RuleActions](Outlook.RuleActions.md)** collection always returns a **RuleAction** object regardless of whether the rule associated with this **RuleActions** collection has defined such a rule action. If the rule has defined and enabled such a rule action, then **[RuleAction.Enabled](Outlook.MoveOrCopyRuleAction.Enabled.md)** will be **True**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]