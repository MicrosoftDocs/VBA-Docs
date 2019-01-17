---
title: Application.WindowDeactivate Event (Publisher)
keywords: vbapb10.chm268435458
f1_keywords:
- vbapb10.chm268435458
ms.prod: publisher
api_name:
- Publisher.Application.WindowDeactivate
ms.assetid: 84473784-7c03-4c9e-3e1b-9bf6ec7e1fbc
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowDeactivate Event (Publisher)

Occurs when the application window is deactivated.


## Syntax

 _expression_. **WindowDeactivate**(**_Wn_**, )

 _expression_ A variable that represents an  **Application** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Wn|Required| **Window**|The window that is being deactivated.|

## Remarks

For information about using events with the Application object, see  [Using Events with the Application Object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

This example minimizes the window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see  [Using Events with the Application Object](../publisher/Concepts/using-events-with-the-application-object-publisher.md)for directions on how to accomplish this.


```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowDeactivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMinimize 
End Sub
```


## See also


 [Application Object](Publisher.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]