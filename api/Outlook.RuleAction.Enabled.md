---
title: RuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2205
f1_keywords:
- vbaol11.chm2205
ms.prod: outlook
api_name:
- Outlook.RuleAction.Enabled
ms.assetid: bea1a0e4-4fad-acc4-0b48-b2f64d996941
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleAction.Enabled property (Outlook)

Returns or sets a **Boolean** that determines if the **[RuleAction](Outlook.RuleAction.md)** is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [RuleAction](Outlook.RuleAction.md) object.


## Remarks

After you enable a rule action, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule action and its enabled state will persist beyond the current session. A rule action is only enabled after it has been saved successfully.

Returns an error if you attempt to enable a rule action that is supported only on a rule of type  **olRuleSend** for a rule of type **olRuleReceive**, or vice versa.


## See also


[RuleAction Object](Outlook.RuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]