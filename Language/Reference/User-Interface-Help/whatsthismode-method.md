---
title: WhatsThisMode method (Visual Basic for Applications)
keywords: vblr6.chm1100685
f1_keywords:
- vblr6.chm1100685
ms.prod: office
api_name:
- Office.WhatsThisMode
ms.assetid: e71fb00c-b323-2b43-94ec-07079e66337f
ms.date: 12/14/2018
localization_priority: Normal
---


# WhatsThisMode method

Causes the mouse pointer to change to the **What's This** pointer and prepares the application to display Help on a selected object. This method exists on the Macintosh, but there is no pointer functionality.

## Syntax

_object_.**WhatsThisMode**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list. If _object_ is omitted, the **[UserForm](userform-window.md)** with the [focus](../../Glossary/vbe-glossary.md#focus) is assumed to be _object_.

## Remarks

Executing the **WhatsThisMode** method places the application in the same state as clicking the **What's This** button on the title bar. The mouse pointer changes to the **What's This** pointer. When the user clicks an object, the **WhatsThisHelpID** property of the clicked object is used to invoke the context-sensitive Help.

## Example

The following example changes the mouse pointer to the **What's This** (question mark) pointer when the user clicks the **UserForm**. If neither the **WhatsThisHelp** or the **WhatsThisButton** property is set to **True** in the Properties window, the following invocation has no effect.


```vb
Private Sub UserForm_Click()
' Turn mouse pointer to What's This question mark.
    WhatsThisMode
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]