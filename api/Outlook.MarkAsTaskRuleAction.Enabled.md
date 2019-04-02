---
title: MarkAsTaskRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2283
f1_keywords:
- vbaol11.chm2283
ms.prod: outlook
api_name:
- Outlook.MarkAsTaskRuleAction.Enabled
ms.assetid: 3e969ccd-7af2-d6db-ab63-d17ce2c2614c
ms.date: 06/08/2017
localization_priority: Normal
---


# MarkAsTaskRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [MarkAsTaskRuleAction](Outlook.MarkAsTaskRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[MarkAsTaskRuleAction Object](Outlook.MarkAsTaskRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]