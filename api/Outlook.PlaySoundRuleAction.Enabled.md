---
title: PlaySoundRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2275
f1_keywords:
- vbaol11.chm2275
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction.Enabled
ms.assetid: 7a8b222e-a9db-f38f-8f8b-a834ff46c39a
ms.date: 06/08/2017
localization_priority: Normal
---


# PlaySoundRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [PlaySoundRuleAction](Outlook.PlaySoundRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[PlaySoundRuleAction Object](Outlook.PlaySoundRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]