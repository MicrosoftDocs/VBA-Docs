---
title: String function (Visual Basic for Applications)
keywords: vblr6.chm1011358
f1_keywords:
- vblr6.chm1011358
ms.prod: office
ms.assetid: d6c5c054-21b9-f777-acae-ac31710ba5c5
ms.date: 12/13/2018
localization_priority: Normal
---


# String function

Returns a **Variant** (**String**) containing a repeating character string of the length specified.

## Syntax

**String**(_number_, _character_)

<br/>

The **String** function syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_number_|Required; [Long](../../Glossary/vbe-glossary.md#long-data-type). Length of the returned string. If _number_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.|
|_character_|Required; [Variant](../../Glossary/vbe-glossary.md#variant-data-type). [Character code](../../Glossary/vbe-glossary.md#character-code) specifying the character or [string expression](../../Glossary/vbe-glossary.md#string-expression) whose first character is used to build the return string. If _character_ contains **Null**, **Null** is returned.|

## Remarks

If you specify a number for _character_ greater than 255, **String** converts the number to a valid character code by using this formula: _character_ **Mod** 256.

## Example

This example uses the **String** function to return repeating character strings of the length specified.


```vb
Dim MyString
MyString = String(5, "*")    ' Returns "*****".
MyString = String(5, 42)    ' Returns "*****".
MyString = String(10, "ABC")    ' Returns "AAAAAAAAAA".

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
