---
title: Spc function (Visual Basic for Applications)
keywords: vblr6.chm1009027
f1_keywords:
- vblr6.chm1009027
ms.prod: office
ms.assetid: a7f2d6f4-6d27-fca5-80ad-648fbf46c002
ms.date: 12/13/2018
localization_priority: Normal
---


# Spc function

Used with the **[Print #](printstatement.md)** statement or the **[Print](print-method.md)** method to position output.

## Syntax

**Spc**(_n_)

The required _n_ [argument](../../Glossary/vbe-glossary.md#argument) is the number of spaces to insert before displaying or printing the next [expression](../../Glossary/vbe-glossary.md#expression) in a list.

## Remarks

If _n_ is less than the output line width, the next print position immediately follows the number of spaces printed. If _n_ is greater than the output line width, **Spc** calculates the next print position using the formula: _currentprintposition_ + (_n_**Mod**_width_).

For example, if the current print position is 24, the output line width is 80, and you specify **Spc**(90), the next print will start at position 34 (current print position + the remainder of 90/80). If the difference between the current print position and the output line width is less than _n_ (or _n_**Mod**_width_), the **Spc** function skips to the beginning of the next line and generates spaces equal to _n_ - (_width_ - _currentprintposition_).

> [!NOTE] 
> Make sure your tabular columns are wide enough to accommodate wide letters.

When you use the **Print** method with a proportionally spaced font, the width of space characters printed by using the **Spc** function is always an average of the width of all characters in the point size for the chosen font. However, there is no correlation between the number of characters printed and the number of fixed-width columns those characters occupy. For example, the uppercase letter W occupies more than one fixed-width column and the lowercase letter i occupies less than one fixed-width column.

## Example

This example uses the **Spc** function to position output in a file and in the [Immediate window](immediate-window.md).

```vb
' The Spc function can be used with the Print # statement.
Open "TESTFILE" For Output As #1    ' Open file for output.
Print #1, "10 spaces between here"; Spc(10); "and here."
Close #1    ' Close file.

```

<br/>

The following statement causes the text to be printed in the Immediate window (by using the **Print** method), preceded by 30 spaces.

```vb
Debug.Print Spc(30); "Thirty spaces later..."

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]