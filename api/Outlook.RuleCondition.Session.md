---
title: RuleCondition.Session property (Outlook)
keywords: vbaol11.chm2327
f1_keywords:
- vbaol11.chm2327
ms.prod: outlook
api_name:
- Outlook.RuleCondition.Session
ms.assetid: bb2163ff-72fb-5712-4618-7dd814b76f9f
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [RuleCondition](Outlook.RuleCondition.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[RuleCondition Object](Outlook.RuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]