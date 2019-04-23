---
title: FromRssFeedRuleCondition.Enabled property (Outlook)
keywords: vbaol11.chm3257
f1_keywords:
- vbaol11.chm3257
ms.prod: outlook
api_name:
- Outlook.FromRssFeedRuleCondition.Enabled
ms.assetid: 162939a7-005b-7762-541c-d7cd2f5e979a
ms.date: 06/08/2017
localization_priority: Normal
---


# FromRssFeedRuleCondition.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [FromRssFeedRuleCondition](Outlook.FromRssFeedRuleCondition.md) object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is enabled only after it has been saved successfully.


## See also


[FromRssFeedRuleCondition Object](Outlook.FromRssFeedRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]