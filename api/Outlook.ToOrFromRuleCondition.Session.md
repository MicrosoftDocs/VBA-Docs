---
title: ToOrFromRuleCondition.Session property (Outlook)
keywords: vbaol11.chm2458
f1_keywords:
- vbaol11.chm2458
ms.prod: outlook
api_name:
- Outlook.ToOrFromRuleCondition.Session
ms.assetid: e2d878c2-ad46-c111-f2e6-9f9af04c1ca5
ms.date: 06/08/2017
localization_priority: Normal
---


# ToOrFromRuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [ToOrFromRuleCondition](Outlook.ToOrFromRuleCondition.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ToOrFromRuleCondition Object](Outlook.ToOrFromRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]