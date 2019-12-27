---
title: Asc function (Visual Basic for Applications)
keywords: vblr6.chm1009247
f1_keywords:
- vblr6.chm1009247
ms.prod: office
ms.assetid: 4c5775f4-792f-f9d0-6eff-41d6fff9048c
ms.date: 12/11/2018
localization_priority: Normal
---


# Asc function

Returns an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) representing the [character code](../../Glossary/vbe-glossary.md#character-code) corresponding to the first letter in a string.

## Syntax

**Asc**(_string_)

The required _string_ [argument](../../Glossary/vbe-glossary.md#argument) is any valid [string expression](../../Glossary/vbe-glossary.md#string-expression). If the _string_ contains no characters, a [run-time error](../../Glossary/vbe-glossary.md#run-time-error) occurs.

## Remarks

The range for returns is 0&ndash;255 on non-DBCS systems, but -32768&ndash;32767 on [DBCS](../../Glossary/vbe-glossary.md#dbcs) systems.

> [!NOTE] 
> The **AscB** function is used with byte data contained in a string. Instead of returning the character code for the first character, **AscB** returns the first byte. The **AscW** function returns the [Unicode](../../Glossary/vbe-glossary.md#unicode) character code except on platforms where Unicode is not supported, in which case, the behavior is identical to the **Asc** function.


> [!NOTE] 
> Visual Basic for the Macintosh does not support Unicode strings. Therefore, **AscW** (_n_) cannot return all Unicode characters for n values in the range of 128&ndash;65,535, as it does in the Windows environment. Instead, **AscW** (_n_) attempts a "best guess" for Unicode values n greater than 127. Therefore, you should not use **AscW** in the Macintosh environment.

The functions **[Chr(), ChrB(), and ChrW()](../chr-function.md)** are the opposite of **Asc(), AscB(), and AscW().** The **Chr()** functions convert an integer to a character string.

## Example

This example uses the **Asc** function to return a character code corresponding to the first letter in the string.


```vb
Dim MyNumber
MyNumber = Asc("A")    ' Returns 65.
MyNumber = Asc("a")    ' Returns 97.
MyNumber = Asc("Apple")    ' Returns 65.

```


## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)
- [Chr(), ChrB(), and ChrW() functions](../chr-function.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
