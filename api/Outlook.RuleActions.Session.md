---
title: RuleActions.Session property (Outlook)
keywords: vbaol11.chm2180
f1_keywords:
- vbaol11.chm2180
ms.prod: outlook
api_name:
- Outlook.RuleActions.Session
ms.assetid: 10b906a5-421c-e858-f8f1-561818425f0a
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]