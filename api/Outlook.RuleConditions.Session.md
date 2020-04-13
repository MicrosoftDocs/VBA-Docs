---
title: RuleConditions.Session property (Outlook)
keywords: vbaol11.chm2298
f1_keywords:
- vbaol11.chm2298
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Session
ms.assetid: 0a214009-1bd1-9631-a80c-e942680ae878
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]