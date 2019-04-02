---
title: MarkAsTaskRuleAction.Session property (Outlook)
keywords: vbaol11.chm2281
f1_keywords:
- vbaol11.chm2281
ms.prod: outlook
api_name:
- Outlook.MarkAsTaskRuleAction.Session
ms.assetid: c98edd5e-e887-4dfe-9e92-1618f506a10b
ms.date: 06/08/2017
localization_priority: Normal
---


# MarkAsTaskRuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [MarkAsTaskRuleAction](Outlook.MarkAsTaskRuleAction.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[MarkAsTaskRuleAction Object](Outlook.MarkAsTaskRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]