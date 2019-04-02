---
title: NewItemAlertRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2292
f1_keywords:
- vbaol11.chm2292
ms.prod: outlook
api_name:
- Outlook.NewItemAlertRuleAction.Enabled
ms.assetid: f3472ffb-ada6-c18d-3953-4a1dd7a25a44
ms.date: 06/08/2017
localization_priority: Normal
---


# NewItemAlertRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [NewItemAlertRuleAction](Outlook.NewItemAlertRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[NewItemAlertRuleAction Object](Outlook.NewItemAlertRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]