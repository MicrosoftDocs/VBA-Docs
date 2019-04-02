---
title: SendRuleAction.Enabled property (Outlook)
keywords: vbaol11.chm2220
f1_keywords:
- vbaol11.chm2220
ms.prod: outlook
api_name:
- Outlook.SendRuleAction.Enabled
ms.assetid: c046cb54-b275-b903-2f9c-dc9a106cdc8a
ms.date: 06/08/2017
localization_priority: Normal
---


# SendRuleAction.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule action is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [SendRuleAction](Outlook.SendRuleAction.md) object.


## Remarks

After you enable a rule, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.


## See also


[SendRuleAction Object](Outlook.SendRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]