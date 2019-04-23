---
title: FormDescription.Session property (Outlook)
keywords: vbaol11.chm181
f1_keywords:
- vbaol11.chm181
ms.prod: outlook
api_name:
- Outlook.FormDescription.Session
ms.assetid: 371dc0ed-f0c6-6c16-930a-f7c5e78b3b54
ms.date: 06/08/2017
localization_priority: Normal
---


# FormDescription.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [FormDescription](Outlook.FormDescription.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[FormDescription Object](Outlook.FormDescription.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]