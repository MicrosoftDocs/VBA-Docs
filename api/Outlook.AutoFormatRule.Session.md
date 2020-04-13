---
title: AutoFormatRule.Session property (Outlook)
keywords: vbaol11.chm2705
f1_keywords:
- vbaol11.chm2705
ms.prod: outlook
api_name:
- Outlook.AutoFormatRule.Session
ms.assetid: b443da40-c6fc-c4a8-c27c-b5f383c8a3ed
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [AutoFormatRule](Outlook.AutoFormatRule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[AutoFormatRule Object](Outlook.AutoFormatRule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]