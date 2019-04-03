---
title: FormNameRuleCondition.Session property (Outlook)
keywords: vbaol11.chm2450
f1_keywords:
- vbaol11.chm2450
ms.prod: outlook
api_name:
- Outlook.FormNameRuleCondition.Session
ms.assetid: ec224a2e-1d45-82d8-0336-9f1f36549b7a
ms.date: 06/08/2017
localization_priority: Normal
---


# FormNameRuleCondition.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [FormNameRuleCondition](Outlook.FormNameRuleCondition.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[FormNameRuleCondition Object](Outlook.FormNameRuleCondition.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]