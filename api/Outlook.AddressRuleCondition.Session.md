---
title: AddressRuleCondition.Session property (Outlook)
keywords: vbaol11.chm2951
f1_keywords:
- vbaol11.chm2951
ms.prod: outlook
api_name:
- Outlook.AddressRuleCondition.Session
ms.assetid: c5134be6-7ce4-dc65-8bde-9c725ef3ba8c
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressRuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AddressRuleCondition](Outlook.AddressRuleCondition.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AddressRuleCondition Object](Outlook.AddressRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]