---
title: AssignToCategoryRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2267
f1_keywords:
- vbaol11.chm2267
ms.prod: outlook
api_name:
- Outlook.AssignToCategoryRuleAction.Enabled
ms.assetid: c6f4558d-fb2a-b732-cfeb-a30f447f0989
ms.date: 06/08/2017
localization_priority: Normal
---


# AssignToCategoryRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents an [AssignToCategoryRuleAction](Outlook.AssignToCategoryRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[AssignToCategoryRuleAction Object](Outlook.AssignToCategoryRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]