---
title: PropertyAccessor.Session property (Outlook)
keywords: vbaol11.chm1968
f1_keywords:
- vbaol11.chm1968
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.Session
ms.assetid: db33aa4e-ad96-2db8-de9d-7aa9dd1a137f
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Remarks

The **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **[NameSpace](Outlook.NameSpace.md)** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]