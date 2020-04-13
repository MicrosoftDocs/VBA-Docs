---
title: AddressRuleCondition.Enabled property (Outlook)
keywords: vbaol11.chm2953
f1_keywords:
- vbaol11.chm2953
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition.Enabled
ms.assetid: 170cd84c-4733-0801-c411-34736e2e1a06
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressRuleCondition.Enabled property (Outlook)

Returns or sets a **Boolean** that determines if the rule condition is enabled. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents an [AddressRuleCondition](Outlook.AddressRuleCondition.md) object.


## Remarks

After you enable a rule condition, you must also save the rule by using  **[Rules.Save](Outlook.Rules.Save.md)** so that the rule condition and its enabled state will persist beyond the current session. A rule condition is only enabled after it have been saved successfully.


## See also


[AddressRuleCondition Object](Outlook.AddressRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]