---
title: CategoryRuleCondition.Session property (Outlook)
keywords: vbaol11.chm2442
f1_keywords:
- vbaol11.chm2442
ms.prod: outlook
api_name:
- Outlook.CategoryRuleCondition.Session
ms.assetid: ee8824ce-0cc8-1e32-1878-721f5e7a81be
ms.date: 06/08/2017
localization_priority: Normal
---


# CategoryRuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [CategoryRuleCondition](Outlook.CategoryRuleCondition.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[CategoryRuleCondition Object](Outlook.CategoryRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]