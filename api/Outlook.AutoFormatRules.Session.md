---
title: AutoFormatRules.Session property (Outlook)
keywords: vbaol11.chm2715
f1_keywords:
- vbaol11.chm2715
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules.Session
ms.assetid: 725f7311-29bd-8536-4625-896cc9baffcb
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRules.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AutoFormatRules](Outlook.AutoFormatRules.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AutoFormatRules Object](Outlook.AutoFormatRules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]