---
title: Assert method (Visual Basic for Applications)
keywords: vblr6.chm1103682
f1_keywords:
- vblr6.chm1103682
ms.prod: office
api_name:
- Office.Assert
ms.assetid: 50bc7f70-d1d0-b23b-e449-f41815cc3178
ms.date: 12/14/2018
localization_priority: Normal
---


# Assert method

Conditionally suspends execution when _booleanexpression_ returns **False** at the line on which the method appears.

## Syntax

_object_.**Assert** _booleanexpression_

<br/>

The **Assert** method syntax has the following object qualifier and argument:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the **[Debug](debug-object.md)** object.|
| _booleanexpression_|Required. An [expression](../../Glossary/vbe-glossary.md#expression) that evaluates to either **True** or **False**.|

## Remarks

**Assert** invocations work only within the [development environment](../../Glossary/vbe-glossary.md#development-environment). When the [module](../../Glossary/vbe-glossary.md#module) is compiled into an executable, the method calls on the **Debug** object are omitted.

All of _booleanexpression_ is always evaluated. For example, even if the first part of an **And** expression evaluates **False**, the entire expression is evaluated.

## Example

The following example shows how to use the **Assert** method. The example requires a form with two button controls on it. The default button names are **Command1** and **Command2**.

When the example runs, clicking the **Command1** button toggles the text on the button between 0 and 1. Clicking **Command2** either does nothing or causes an assertion, depending on the value displayed on **Command1**. The assertion stops execution with the last statement executed, the Debug.Assert line, highlighted.

```vb
Option Explicit
Private blnAssert As Boolean
Private intNumber As Integer

Private Sub Command1_Click()
    blnAssert = Not blnAssert
    intNumber = IIf(intNumber <> 0, 0, 1)
    Command1.Caption = intNumber
End Sub

Private Sub Command2_Click()
    Debug.Assert blnAssert
End Sub

Private Sub Form_Load()
    Command1.Caption = intNumber
    Command2.Caption = "Assert Tester"
End Sub
```

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
