---
title: TextRuleCondition.Enabled property (Outlook)
keywords: vbaol11.chm2476
f1_keywords:
- vbaol11.chm2476
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition.Enabled
ms.assetid: 7027c22b-08fa-d1b0-f664-8c4a26722cbb
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRuleCondition.Enabled property (Outlook)

Returns or sets a  **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [TextRuleCondition](Outlook.TextRuleCondition.md) object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it have been saved successfully.

Returns an error if you attempt to enable a rule condition that is supported only on a rule of type  **olRuleSend** for a rule of type **olRuleReceive**, or vice versa. For more information on suppport by rules for receiving messages or rules for sending messages, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## See also


[TextRuleCondition Object](Outlook.TextRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]