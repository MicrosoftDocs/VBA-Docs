---
title: Resize event (Visual Basic for Applications)
keywords: vblr6.chm1107498
f1_keywords:
- vblr6.chm1107498
ms.prod: office
api_name:
- Office.Resize
ms.assetid: d7ea6a67-1d51-0dee-0b23-19ca748557ea
ms.date: 12/11/2018
localization_priority: Normal
---


# Resize event

Occurs when a user form is resized.

## Syntax

**Private Sub UserForm_Resize( )**

## Remarks

Use a **Resize** event [procedure](../../Glossary/vbe-glossary.md#procedure) to move or resize [controls](../../Glossary/vbe-glossary.md#control) when the parent **[UserForm](userform-window.md)** is resized. You can also use this event procedure to recalculate [variables](../../Glossary/vbe-glossary.md#variable) or [properties](../../Glossary/vbe-glossary.md#property).

## Example

The following example uses the **Activate** and **Click** events to illustrate triggering of the **UserForm** **Resize** event. As the user clicks the client area of the form, it grows or shrinks, and the new height is specified in the title bar. Note that the **[Tag](tag-property.md)** property is used to store the **UserForm**'s initial height.


```vb
' Activate event for UserForm1
Private Sub UserForm_Activate()
    UserForm1.Caption = "Click me to make me taller!"
    Tag = Height    ' Save the initial height.
End Sub

' Click event for UserForm1
Private Sub UserForm_Click()
    Dim NewHeight As Single
    NewHeight = Height
    ' If the form is small, make it tall.
    If NewHeight = Val(Tag) Then
        Height = Val(Tag) * 2
    Else
    ' If the form is tall, make it small.
        Height = Val(Tag)
    End If
End Sub

' Resize event for UserForm1
Private Sub UserForm_Resize()
    UserForm1.Caption = "New Height: " & Height & "  " & "Click to resize me!"
End Sub
```

## See also

- [Events (Visual Basic Add-In Model)](../visual-basic-add-in-model/events-visual-basic-add-in-model.md)
- [Events (Visual Basic for Applications)](../events-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]