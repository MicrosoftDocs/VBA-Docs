---
title: Application.Version property (Publisher)
keywords: vbapb10.chm131121
f1_keywords:
- vbapb10.chm131121
ms.prod: publisher
api_name:
- Publisher.Application.Version
ms.assetid: ffec5bca-cd81-77c6-d80b-e629abfa6dec
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.Version property (Publisher)

Returns a **String** indicating the version number of the currently-installed copy of Microsoft Publisher. Read-only.


## Syntax

_expression_.**Version**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

String


## Example

The following example displays the version and build number of the currently-installed copy of Publisher.

```vb
MsgBox "You are currently running Microsoft Publisher, " _ 
 & " version " & Application.Version & ", build " _ 
 & Application.Build & "." 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]