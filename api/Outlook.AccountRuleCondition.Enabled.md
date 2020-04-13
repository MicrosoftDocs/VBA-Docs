---
title: AccountRuleCondition.Enabled property (Outlook)
keywords: vbaol11.chm2381
f1_keywords:
- vbaol11.chm2381
ms.prod: outlook
api_name:
- Outlook.AccountRuleCondition.Enabled
ms.assetid: 834b45ee-f140-7e02-47ea-00e68ae6580c
ms.date: 06/08/2017
localization_priority: Normal
---


# AccountRuleCondition.Enabled property (Outlook)

Returns or sets a **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents an [AccountRuleCondition](Outlook.AccountRuleCondition.md) object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it has been saved successfully.


## See also


[AccountRuleCondition Object](Outlook.AccountRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]