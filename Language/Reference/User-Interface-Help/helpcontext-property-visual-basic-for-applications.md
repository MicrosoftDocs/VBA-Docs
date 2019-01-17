---
title: HelpContext property (Visual Basic for Applications)
keywords: vblr6.chm1014190
f1_keywords:
- vblr6.chm1014190
ms.prod: office
ms.assetid: 5cfd1f6c-1d91-623c-dbb0-3431d5837881
ms.date: 12/19/2018
localization_priority: Normal
---


# HelpContext property

Returns or sets a [string expression](../../Glossary/vbe-glossary.md#string-expression) containing the context ID for a topic in a Help file. Read/write.

## Remarks

The **HelpContext** [property](../../Glossary/vbe-glossary.md#property) is used to automatically display the Help topic specified in the **[HelpFile](helpfile-property.md)** property. 

If both **HelpFile** and **HelpContext** are empty, the value of **[Number](number-property-visual-basic-for-applications.md)** is checked. If **Number** corresponds to a Visual Basic [run-time error](../../Glossary/vbe-glossary.md#run-time-error) value, the Visual Basic Help context ID for the error is used. If the **Number** value doesn't correspond to a Visual Basic error, the contents screen for the Visual Basic Help file is displayed.

> [!NOTE] 
> You should write routines in your application to handle typical errors. When programming with an object, you can use the object's Help file to improve the quality of your error handling, or to display a meaningful message to your user if the error isn't recoverable.


## Example

This example uses the **HelpContext** property of the **[Err](err-object.md)** object to show the Visual Basic Help topic for the `Overflow` error.


```vb
Dim Msg
Err.Clear
On Error Resume Next
Err.Raise 6 ' Generate "Overflow" error.
If Err.Number <> 0 Then
    Msg = "Press F1 or HELP to see " & Err.HelpFile & " topic for" & _
    " the following HelpContext: " & Err. HelpContext
    MsgBox Msg, , "Error: " & Err.Description, Err.HelpFile, _
Err.HelpContext
End If
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]