---
title: TextRuleCondition.ConditionType property (Outlook)
keywords: vbaol11.chm2477
f1_keywords:
- vbaol11.chm2477
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition.ConditionType
ms.assetid: 2dbc7979-deae-fbb8-9def-8c906657024a
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRuleCondition.ConditionType property (Outlook)

Returns a constant from the  **[OlRuleConditionType](Outlook.OlRuleConditionType.md)** enumeration that indicates the type of rule condition. Read-only.


## Syntax

_expression_. `ConditionType`

_expression_ A variable that represents a [TextRuleCondition](Outlook.TextRuleCondition.md) object.


## Remarks

The value of  **ConditionType** depends on the type of rule condition, as several types of rule conditions use the **[TextRuleCondition](Outlook.TextRuleCondition.md)** object: **olConditionBody**, **olConditionBodyOrSubject**, **olConditionMessageHeader**, and **olConditionSubject**. Except for **olConditionMessageHeader**, which is supported only by rules for receiving messages, all these types of conditions are supported by rules for receiving messages as well as rules for sending messages. For more information, see [Specify Rule Conditions](../outlook/How-to/Rules/specifying-rule-conditions.md).


## See also


[TextRuleCondition Object](Outlook.TextRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]