---
title: Rule.Session property (Outlook)
keywords: vbaol11.chm2166
f1_keywords:
- vbaol11.chm2166
ms.prod: outlook
api_name:
- Outlook.Rule.Session
ms.assetid: 7502f919-cf8f-d795-87b1-9812c0d150d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Rule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Rule](Outlook.Rule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Rule Object](Outlook.Rule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]