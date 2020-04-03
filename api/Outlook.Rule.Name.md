---
title: Rule.Name property (Outlook)
keywords: vbaol11.chm2168
f1_keywords:
- vbaol11.chm2168
ms.prod: outlook
api_name:
- Outlook.Rule.Name
ms.assetid: 6c559ffe-b25c-ff49-31d1-1fd44935a8f3
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule.Name property (Outlook)

Returns or sets a **String** representing the name of the rule. Read/write.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Rule](Outlook.Rule.md) object.


## Remarks

 **Name** is the default property and an indexer for a rule in a **[Rules](Outlook.Rules.md)** collection object. It corresponds to **PidTagRuleMsgName** (0x65EC001E).

Rule names are not unique among rules in the same collection.


## See also


[Rule Object](Outlook.Rule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]