---
title: Application.WindowActivate Event (Publisher)
keywords: vbapb10.chm268435457
f1_keywords:
- vbapb10.chm268435457
ms.prod: publisher
api_name:
- Publisher.Application.WindowActivate
ms.assetid: a7e4e396-9661-763c-8e41-dc279757af94
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowActivate Event (Publisher)

Occurs when the application window is activated.


## Syntax

 _expression_. **WindowActivate**(**_Wn_**, )

 _expression_ A variable that represents an  **Application** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Wn|Required| **Window**|The window that is being activated.|

## Remarks

For information about using events with the Application object, see  [Using Events with the Application Object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

This example maximizes the Microsoft Publisher window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see  [Using Events with the Application Object](../publisher/Concepts/using-events-with-the-application-object-publisher.md)for directions on how to accomplish this.


```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowActivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMaximize 
End Sub
```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]