---
title: Application.Screen property (Access)
keywords: vbaac10.chm12510
f1_keywords:
- vbaac10.chm12510
ms.prod: access
api_name:
- Access.Application.Screen
ms.assetid: d6faa33a-7701-d270-3bc7-04d53ac9303a
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Screen property (Access)

You can use the **Screen** property to return a reference the **[Screen](Access.Screen.md)** object and its related properties. Read-only.


## Syntax

_expression_.**Screen**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

Use the **Screen** object to refer to a particular form, report, or control that has the focus.


## Example

The following example demonstrates how to change the cursor to an hourglass and back again to signify that some background activity is occurring.


```vb
Application.Screen.MousePointer = 11 ' Hourglass' Do some background activity.Application.Screen.MousePointer = 0 ' Back to normal
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]