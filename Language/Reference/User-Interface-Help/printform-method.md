---
title: PrintForm method (Visual Basic for Applications)
keywords: vblr6.chm916130
f1_keywords:
- vblr6.chm916130
ms.prod: office
api_name:
- Office.PrintForm
ms.assetid: d4481074-6ecf-b845-2a51-ef34dcdc82ab
ms.date: 12/14/2018
localization_priority: Normal
---


# PrintForm method

Sends a bit-by-bit image of a **[UserForm](userform-window.md)** object to the printer.

## Syntax

_object_.**PrintForm**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list. If _object_ is omitted, the **UserForm** with the [focus](../../Glossary/vbe-glossary.md#focus) is assumed to be _object_.

## Remarks

**PrintForm** prints all visible objects and [bitmaps](../../Glossary/vbe-glossary.md#bitmap) of the **UserForm** object. **PrintForm** also prints graphics added to a **UserForm** object.

The printer used by **PrintForm** is determined by the operating system's **Control Panel** settings.

## Example

In the following example, the client area of the form is printed when the user clicks the form.

```vb
' This is the click event for UserForm1
Private Sub UserForm_Click()
    UserForm1.PrintForm
End Sub
```


## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]