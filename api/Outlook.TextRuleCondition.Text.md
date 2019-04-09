---
title: TextRuleCondition.Text property (Outlook)
keywords: vbaol11.chm2478
f1_keywords:
- vbaol11.chm2478
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition.Text
ms.assetid: 615f47e9-2c43-a473-33f6-46765ccd3903
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRuleCondition.Text property (Outlook)

Returns or sets an array of  **String** elements that represents the text to be evaluated by the rule condition. Read/write.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a [TextRuleCondition](Outlook.TextRuleCondition.md) object.


## Remarks

You can assign an array with one string or multiple strings for evaluation. Multiple text strings assigned in an array are evaluated using the logical OR operation.


## See also


[TextRuleCondition Object](Outlook.TextRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]