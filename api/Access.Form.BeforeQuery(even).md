---
title: Form.BeforeQuery event (Access)
keywords: vbaac10.chm13671
f1_keywords:
- vbaac10.chm13671
ms.prod: access
api_name:
- Access.Form.BeforeQuery
ms.assetid: 07d9ba3f-25dc-f448-5c99-8c1e4ca5ab20
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.BeforeQuery event (Access)

Occurs when the specified PivotTable view queries its data source.


## Syntax

_expression_.**BeforeQuery**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Return value

Nothing


## Remarks

This event occurs quite frequently. Some examples of actions that trigger this event include adding fields to the PivotTable view, moving fields, sorting, or filtering data.


## Example

The following example demonstrates the syntax for a subroutine that traps the **BeforeQuery** event.

```vb
Private Sub Form_BeforeQuery() 
 MsgBox "The PivotTable view is about to query its data source." 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]