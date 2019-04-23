---
title: AssignToCategoryRuleAction.Session property (Outlook)
keywords: vbaol11.chm2265
f1_keywords:
- vbaol11.chm2265
ms.prod: outlook
api_name:
- Outlook.AssignToCategoryRuleAction.Session
ms.assetid: 4ee91dde-9f5d-101f-f259-98192e45a76d
ms.date: 06/08/2017
localization_priority: Normal
---


# AssignToCategoryRuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AssignToCategoryRuleAction](Outlook.AssignToCategoryRuleAction.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AssignToCategoryRuleAction Object](Outlook.AssignToCategoryRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]