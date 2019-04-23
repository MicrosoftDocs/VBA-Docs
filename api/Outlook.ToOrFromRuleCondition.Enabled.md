---
title: ToOrFromRuleCondition.Enabled property (Outlook)
keywords: vbaol11.chm2460
f1_keywords:
- vbaol11.chm2460
ms.prod: outlook
api_name:
- Outlook.ToOrFromRuleCondition.Enabled
ms.assetid: 31e43906-b47a-95e3-d51b-3fa6af553fad
ms.date: 06/08/2017
localization_priority: Normal
---


# ToOrFromRuleCondition.Enabled property (Outlook)

Returns a  **Boolean** value that indicates whether the rule condition is enabled. Read/write


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a [ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md) object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it have been saved successfully.

Returns an error if you attempt to enable a rule condition that is supported only on a rule of type  **olRuleSend** for a rule of type **olRuleReceive**, or vice versa. For more information on suppport by rules for receiving messages or rules for sending messages, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## See also


[ToOrFromRuleCondition Object](Outlook.ToOrFromRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]