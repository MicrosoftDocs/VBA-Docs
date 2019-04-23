---
title: Unload statement (VBA)
keywords: vblr6.chm1100684
f1_keywords:
- vblr6.chm1100684
ms.prod: office
ms.assetid: 5fa03dfb-686d-b266-18ba-e4c50afd63ea
ms.date: 12/03/2018
localization_priority: Normal
---


# Unload statement

Removes an object from memory.

## Syntax

**Unload** _object_

The required _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

When an object is unloaded, it's removed from memory and all memory associated with the object is reclaimed. Until it is placed in memory again by using the **[Load](load-statement.md)** statement, a user can't interact with an object, and the object can't be manipulated programmatically.

## Example

The following example assumes two **UserForms** in a program. In UserForm1's **Initialize** event, UserForm2 is loaded and shown. When the user clicks UserForm2, it is unloaded and UserForm1 appears. When UserForm1 is clicked, it is unloaded in turn.


```vb
' This is the Initialize event procedure for UserForm1 
Private Sub UserForm_Initialize() 
 Load UserForm2 
 UserForm2.Show 
End Sub 
' This is the Click event for UserForm2 
Private Sub UserForm_Click() 
 Unload UserForm2 
End Sub 
 
' This is the Click event for UserForm1 
Private Sub UserForm_Click() 
 Unload UserForm1 
End Sub
```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]