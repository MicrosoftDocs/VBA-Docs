---
title: Form.OnDisconnect event (Access)
keywords: vbaac10.chm13668
f1_keywords:
- vbaac10.chm13668
ms.prod: access
api_name:
- Access.Form.OnDisconnect
ms.assetid: b5b2a18b-d159-c122-c35e-fe749d755f0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Form.OnDisconnect event (Access)

Occurs when the specified PivotTable view disconnects from a data source.


## Syntax

_expression_. `OnDisconnect`

_expression_ A variable that represents a [Form](Access.Form.md) object.


## Return value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **OnDisconnect** event.


```vb
Private Sub Form_OnDisconnect() 
 MsgBox "The PivotTable View has " _ 
 & "disconnected from its data source!" 
End Sub
```


## See also


[Form Object](Access.Form.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]