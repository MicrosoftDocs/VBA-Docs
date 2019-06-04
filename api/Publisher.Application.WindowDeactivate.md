---
title: Application.WindowDeactivate event (Publisher)
keywords: vbapb10.chm268435458
f1_keywords:
- vbapb10.chm268435458
ms.prod: publisher
api_name:
- Publisher.Application.WindowDeactivate
ms.assetid: 84473784-7c03-4c9e-3e1b-9bf6ec7e1fbc
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.WindowDeactivate event (Publisher)

Occurs when the application window is deactivated.


## Syntax

_expression_.**WindowDeactivate** (_Wn_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Wn_|Required| **Window**|The window that is being deactivated.|


## Example

This example minimizes the Microsoft Publisher window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work. For directions about how to accomplish this, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md). 


```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowDeactivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMinimize 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]