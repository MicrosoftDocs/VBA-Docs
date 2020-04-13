---
title: Rules.Session property (Outlook)
keywords: vbaol11.chm2156
f1_keywords:
- vbaol11.chm2156
ms.prod: outlook
api_name:
- Outlook.Rules.Session
ms.assetid: c544e009-623c-3e4d-b71a-9177dcfcc668
ms.date: 06/08/2017
localization_priority: Normal
---


# Rules.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Rules](Outlook.Rules.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Rules Object](Outlook.Rules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]