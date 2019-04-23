---
title: SendRuleAction.Session property (Outlook)
keywords: vbaol11.chm2218
f1_keywords:
- vbaol11.chm2218
ms.prod: outlook
api_name:
- Outlook.SendRuleAction.Session
ms.assetid: 0d0b9289-0381-fe88-d4e7-1d0197ce6d6b
ms.date: 06/08/2017
localization_priority: Normal
---


# SendRuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [SendRuleAction](Outlook.SendRuleAction.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[SendRuleAction Object](Outlook.SendRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]