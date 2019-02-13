---
title: Form.OnConnect event (Access)
keywords: vbaac10.chm13667
f1_keywords:
- vbaac10.chm13667
ms.prod: access
api_name:
- Access.Form.OnConnect
ms.assetid: 39966052-0e06-bde9-142f-ee74d16a9973
ms.date: 06/08/2017
localization_priority: Normal
---


# Form.OnConnect event (Access)

Occurs when the specified PivotTable view connects to a data source.


## Syntax

_expression_. `OnConnect`

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Return value

nothing


## Example

The following example demonstrates the syntax for a subroutine that traps the  **OnConnect** event.


```vb
Private Sub Form_OnConnect() 
 MsgBox "The PivotTable View has " _ 
 & "connected to its data source!" 
End Sub
```


## See also


[Form Object](Access.Form.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]