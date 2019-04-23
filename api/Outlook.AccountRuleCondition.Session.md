---
title: AccountRuleCondition.Session property (Outlook)
keywords: vbaol11.chm2379
f1_keywords:
- vbaol11.chm2379
ms.prod: outlook
api_name:
- Outlook.AccountRuleCondition.Session
ms.assetid: 1bcc0f04-a3a1-40e5-5853-938e284db89f
ms.date: 06/08/2017
localization_priority: Normal
---


# AccountRuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AccountRuleCondition](Outlook.AccountRuleCondition.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AccountRuleCondition Object](Outlook.AccountRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]