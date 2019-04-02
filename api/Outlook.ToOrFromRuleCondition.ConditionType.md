---
title: ToOrFromRuleCondition.ConditionType property (Outlook)
keywords: vbaol11.chm2461
f1_keywords:
- vbaol11.chm2461
ms.prod: outlook
api_name:
- Outlook.ToOrFromRuleCondition.ConditionType
ms.assetid: a5c6e08c-643e-965d-cd3e-b434f20579a0
ms.date: 06/08/2017
localization_priority: Normal
---


# ToOrFromRuleCondition.ConditionType property (Outlook)

Returns a constant from the  **[OlRuleConditionType](Outlook.OlRuleConditionType.md)** enumeration that indicates the type of rule condition. Read-only.


## Syntax

_expression_. `ConditionType`

_expression_ A variable that represents a [ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md) object.


## Remarks

 **ConditionType** depends on the type of rule condition, as two types of rule conditions use the **[ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md)** object: **olConditionFrom** and **olConditionSentTo**. **olConditionFrom** is supported only by rules for receiving messages, while **olConditionSentTo** is supported by rules for receiving messages as well as rules for sending messages. For more information, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## See also


[ToOrFromRuleCondition Object](Outlook.ToOrFromRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]