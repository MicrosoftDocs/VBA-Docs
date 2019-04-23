---
title: Terminate event (Visual Basic for Applications)
keywords: vblr6.chm1107499
f1_keywords:
- vblr6.chm1107499
ms.prod: office
ms.assetid: f386e522-fc8a-f073-668d-e804dca9de49
ms.date: 12/11/2018
localization_priority: Normal
---


# Terminate event

Occurs when all references to an instance of an object are removed from memory by setting all [variables](../../Glossary/vbe-glossary.md#variable) that refer to the object to **[Nothing](nothing-keyword.md)** or when the last reference to the object goes out of [scope](../../Glossary/vbe-glossary.md#scope).

## Syntax

**Private Sub** _object_**_Terminate( )**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Remarks

The **Terminate** event occurs after the object is unloaded. The **Terminate** event isn't triggered if the instances of the **[UserForm](userform-window.md)** or [class](../../Glossary/vbe-glossary.md#class) are removed from memory because the application terminated abnormally. 

For example, if your application invokes the **[End](end-statement.md)** statement before removing all existing instances of the class or **UserForm** from memory, the **Terminate** event isn't triggered for that class or **UserForm**.

## Example

The following event procedures cause a **UserForm** to beep for a few seconds after the user clicks the client area to dismiss the form.

```vb
Private Sub UserForm_Activate()
    UserForm1.Caption = "Click me to kill me!"
End Sub

Private Sub UserForm_Click()
  Unload Me
End Sub

Private Sub UserForm_Terminate()
    Dim Count As Integer
    For Count = 1 To 100
        Beep
    Next
End Sub
```


## See also

- [Events (Visual Basic Add-In Model)](../visual-basic-add-in-model/events-visual-basic-add-in-model.md)
- [Events (Visual Basic for Applications)](../events-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
