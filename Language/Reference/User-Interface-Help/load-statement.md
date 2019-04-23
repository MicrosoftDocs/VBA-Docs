---
title: Load statement (VBA)
keywords: vblr6.chm1100680
f1_keywords:
- vblr6.chm1100680
ms.prod: office
ms.assetid: 58e13f8f-3a3b-99d1-bf05-575ddf42c7c7
ms.date: 12/03/2018
localization_priority: Normal
---


# Load statement

Loads an object but doesn't show it.

## Syntax

**Load** _object_

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies** To list.

## Remarks

When an object is loaded, it is placed in memory, but isn't visible. Use the **[Show](show-method.md)** method to make the object visible. Until an object is visible, a user can't interact with it. The object can be manipulated programmatically in its **[Initialize](initialize-event.md)** event procedure.

## Example

In the following example, UserForm2 is loaded during UserForm1's **Initialize** event. Subsequent clicking on UserForm2 reveals UserForm1.


```vb
' This is the Initialize event procedure for UserForm1 
Private Sub UserForm_Initialize() 
 Load UserForm2 
 UserForm2.Show 
End Sub 
' This is the Click event of UserForm2 
Private Sub UserForm_Click() 
 UserForm2.Hide 
End Sub 
 
' This is the click event for UserForm1 
Private Sub UserForm_Click() 
 UserForm2.Show 
End Sub
```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]