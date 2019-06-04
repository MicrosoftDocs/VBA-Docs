---
title: Application.WindowActivate event (Publisher)
keywords: vbapb10.chm268435457
f1_keywords:
- vbapb10.chm268435457
ms.prod: publisher
api_name:
- Publisher.Application.WindowActivate
ms.assetid: a7e4e396-9661-763c-8e41-dc279757af94
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.WindowActivate event (Publisher)

Occurs when the application window is activated.


## Syntax

_expression_.**WindowActivate** (_Wn_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Wn_|Required| **Window**|The window that is being activated.|


## Example

This example maximizes the Microsoft Publisher window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work. For directions about how to accomplish this, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md). 

```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowActivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMaximize 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]