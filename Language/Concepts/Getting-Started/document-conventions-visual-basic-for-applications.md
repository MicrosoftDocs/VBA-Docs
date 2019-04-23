---
title: Document conventions (VBA)
ms.prod: office
ms.assetid: 1eece8df-7e11-f66d-a2b7-18985c288e81
ms.date: 12/21/2018
localization_priority: Normal
---


# Document conventions (VBA)

## Typographic conventions

Visual Basic documentation uses the following typographic conventions.

|Convention|Description|
|:-----|:-----|
|**Sub**, **If**, **ChDir**, **Print**, **True**, **Debug**|Words in bold with initial letter capitalized indicate language-specific keywords.|
|**Setup**|Words you are instructed to type appear in bold.|
| _object_, _varname_, _arglist_|Italic, lowercase letters indicate placeholders for information you supply.|
|**_pathname_**, **_filenumber_**|Bold, italic, and lowercase letters indicate placeholders for arguments where you can use either positional or [named-argument](../../Glossary/vbe-glossary.md#named-argument) syntax.|
|[ _expressionlist_ ]|In syntax, items inside brackets are optional.|
|`{While | Until}`|In syntax, braces and a vertical bar indicate a mandatory choice between two or more items. You must choose one of the items unless all of the items are also enclosed in brackets. For example: `[{This | That}]`|
|ESC, ENTER|Words in capital letters indicate key names and key sequences.|
|ALT+F1, CTRL+R|A plus sign (+) between key names indicates a combination of keys. For example, ALT+F1 means hold down the ALT key while pressing the F1 key.|

## Code conventions

The following code conventions are used.

This font is used for code, variables, and error message text.

```vb
Sub HelloButton_Click()
Readout.Text = _
"Hello, world!"
End Sub
```

<br/>

An apostrophe (') introduces code comments.

```vb
' This is a comment; these two lines
' are ignored when the program is running.
```

<br/>

Lines too long to fit on one line (except comments) may be continued on the next line by using a line-continuation character, which is a single leading space followed by an underscore ( _):

```vb
Sub Form_MouseDown (Button As Integer, _
Shift As Integer, X As Single, Y As Single)
```

## See also

- [Visual Basic naming rules](visual-basic-naming-rules.md)
- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]