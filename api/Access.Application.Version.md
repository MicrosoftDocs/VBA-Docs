---
title: Application.Version property (Access)
keywords: vbaac10.chm12599
f1_keywords:
- vbaac10.chm12599
ms.prod: access
api_name:
- Access.Application.Version
ms.assetid: 3fd0113f-5c8f-0477-6030-cf548f7cb2ff
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.Version property (Access)

Returns a **String** indicating the version number of the currently installed copy of Microsoft Access. Read-only.


## Syntax

_expression_.**Version**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Example

The following example displays the version and build number of the currently-installed copy of Microsoft Access.


```vb
MsgBox "You are currently running Microsoft Access, " _ 
 & " version " & Application.Version & ", build " _ 
 & Application.Build & "." 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]