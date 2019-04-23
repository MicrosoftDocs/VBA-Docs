---
title: Initialize event (Visual Basic for Applications)
keywords: vblr6.chm916308
f1_keywords:
- vblr6.chm916308
ms.prod: office
api_name:
- Office.Initialize
ms.assetid: b6405bb0-21f6-2654-010b-2a14b418c43d
ms.date: 12/11/2018
localization_priority: Normal
---


# Initialize event

Occurs after an object is loaded, but before it's shown.

## Syntax

**Private Sub** _object_**_Initialize( )**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

The **Initialize** event is typically used to prepare an application or **[UserForm](userform-window.md)** for use.[Variables](../../Glossary/vbe-glossary.md#variable) are assigned initial values, and controls may be moved or resized to accommodate initialization data.

## Example

The following example assumes two **UserForms** in a program. In UserForm1's **Initialize** event, UserForm2 is loaded and shown. When the user clicks UserForm2, it is hidden and UserForm1 appears. When UserForm1 is clicked, UserForm2 is shown again.


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

- [Events (Visual Basic Add-In Model)](../visual-basic-add-in-model/events-visual-basic-add-in-model.md)
- [Events (Visual Basic for Applications)](../events-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
