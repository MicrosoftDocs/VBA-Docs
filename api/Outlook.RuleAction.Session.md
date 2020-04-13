---
title: RuleAction.Session property (Outlook)
keywords: vbaol11.chm2203
f1_keywords:
- vbaol11.chm2203
ms.prod: outlook
api_name:
- Outlook.RuleAction.Session
ms.assetid: a80c6148-0eb0-19c0-4d3e-a3a535624773
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [RuleAction](Outlook.RuleAction.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[RuleAction Object](Outlook.RuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]