---
title: Columns.Session property (Outlook)
keywords: vbaol11.chm2737
f1_keywords:
- vbaol11.chm2737
ms.prod: outlook
api_name:
- Outlook.Columns.Session
ms.assetid: 999b39d6-ed92-021c-ed29-96227f91fce3
ms.date: 06/08/2017
localization_priority: Normal
---


# Columns.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Columns](Outlook.Columns.md) object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Columns Object](Outlook.Columns.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]