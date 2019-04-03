---
title: NewItemAlertRuleAction.Session property (Outlook)
keywords: vbaol11.chm2290
f1_keywords:
- vbaol11.chm2290
ms.prod: outlook
api_name:
- Outlook.NewItemAlertRuleAction.Session
ms.assetid: 7ea1ec54-ccf3-2626-f709-4f9ba54d80a3
ms.date: 06/08/2017
localization_priority: Normal
---


# NewItemAlertRuleAction.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NewItemAlertRuleAction](Outlook.NewItemAlertRuleAction.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NewItemAlertRuleAction Object](Outlook.NewItemAlertRuleAction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]