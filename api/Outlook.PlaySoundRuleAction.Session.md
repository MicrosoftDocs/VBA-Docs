---
title: PlaySoundRuleAction.Session property (Outlook)
keywords: vbaol11.chm2273
f1_keywords:
- vbaol11.chm2273
ms.prod: outlook
api_name:
- Outlook.PlaySoundRuleAction.Session
ms.assetid: 8d3e9f6e-848d-9879-61a8-7662858674d4
ms.date: 06/08/2017
localization_priority: Normal
---


# PlaySoundRuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [PlaySoundRuleAction](Outlook.PlaySoundRuleAction.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[PlaySoundRuleAction Object](Outlook.PlaySoundRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]