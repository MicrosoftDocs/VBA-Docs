---
title: MoveOrCopyRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2212
f1_keywords:
- vbaol11.chm2212
ms.prod: outlook
api_name:
- Outlook.MoveOrCopyRuleAction.Enabled
ms.assetid: 795374af-a8de-b771-97df-3d9e82949af0
ms.date: 06/08/2017
localization_priority: Normal
---


# MoveOrCopyRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [MoveOrCopyRuleAction](Outlook.MoveOrCopyRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[MoveOrCopyRuleAction Object](Outlook.MoveOrCopyRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]