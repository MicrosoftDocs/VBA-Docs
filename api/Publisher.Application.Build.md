---
title: Application.Build property (Publisher)
keywords: vbapb10.chm131078
f1_keywords:
- vbapb10.chm131078
ms.prod: publisher
api_name:
- Publisher.Application.Build
ms.assetid: e0d4bb8e-5185-3d3c-fd80-c1e3c3902b2c
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.Build property (Publisher)

Returns a **Long** that represents the Microsoft Publisher build number. Read-only.


## Syntax

_expression_.**Build**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

Long


## Example

This example displays the Publisher build number.

```vb
Sub BuildNumber() 
 MsgBox Prompt:="The current Microsoft Publisher build number is " & _ 
 Application.Build, Title:="Microsoft Publisher Build" 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]