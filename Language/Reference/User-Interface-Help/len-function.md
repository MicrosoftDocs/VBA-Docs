---
title: Len function (Visual Basic for Applications)
keywords: vblr6.chm1011065
f1_keywords:
- vblr6.chm1011065
ms.prod: office
ms.assetid: 5b5b8789-90cc-ac2c-e6a7-1da1d684bd81
ms.date: 12/13/2018
localization_priority: Normal
---


# Len function

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the number of characters in a string or the number of bytes required to store a [variable](../../Glossary/vbe-glossary.md#variable).

## Syntax

**Len**(_string_ | _varname_)

<br/>

The **Len** function syntax has these parts:

|Part|Description|
|:-----|:-----|
| _string_|Any valid [string expression](../../Glossary/vbe-glossary.md#string-expression). If _string_ contains [Null](../../Glossary/vbe-glossary.md#null), **Null** is returned.|
| _varname_|Any valid [variable](../../Glossary/vbe-glossary.md#variable) name. If _varname_ contains **Null**, **Null** is returned. If _varname_ is a [Variant](../../Glossary/vbe-glossary.md#variant-data-type), **Len** treats it the same as a **String** and always returns the number of characters it contains.|

## Remarks

One (and only one) of the two possible [arguments](../../Glossary/vbe-glossary.md#argument) must be specified. With [user-defined types](../../Glossary/vbe-glossary.md#user-defined-type), **Len** returns the size as it will be written to the file.

> [!NOTE] 
> Use the **LenB** function with byte data contained in a string, as in double-byte character set (DBCS) languages. Instead of returning the number of characters in a string, **LenB** returns the number of bytes used to represent that string. With user-defined types, **LenB** returns the in-memory size, including any padding between elements. For sample code that uses **LenB**, see the second example in the example topic.

> [!NOTE] 
> **Len** may not be able to determine the actual number of storage bytes required when used with variable-length strings in user-defined [data types](../../Glossary/vbe-glossary.md#data-type).


## Example

The first example uses **Len** to return the number of characters in a string or the number of bytes required to store a variable. The **Type...End Type** block defining `CustomerRecord` must be preceded by the keyword **Private** if it appears in a class module. In a standard module, a **Type** statement can be **Public**.


```vb
Type CustomerRecord    ' Define user-defined type.
    ID As Integer    ' Place this definition in a 
    Name As String * 10    ' standard module.
    Address As String * 30
End Type

Dim Customer As CustomerRecord    ' Declare variables.
Dim MyInt As Integer, MyCur As Currency
Dim MyString, MyLen
MyString = "Hello World"    ' Initialize variable.
MyLen = Len(MyInt)    ' Returns 2.
MyLen = Len(Customer)    ' Returns 42.
MyLen = Len(MyString)    ' Returns 11.
MyLen = Len(MyCur)    ' Returns 8.

```

<br/>

The second example uses **LenB** and a user-defined function (**LenMbcs**) to return the number of byte characters in a string if ANSI is used to represent the string.

```vb
Function LenMbcs (ByVal str as String)
    LenMbcs = LenB(StrConv(str, vbFromUnicode))
End Function

Dim MyString, MyLen
MyString = "ABc"
' Where "A" and "B" are DBCS and "c" is SBCS.
MyLen = Len(MyString)
' Returns 3 - 3 characters in the string.
MyLen = LenB(MyString)
' Returns 6 - 6 bytes used for Unicode.
MyLen = LenMbcs(MyString)
' Returns 5 - 5 bytes used for ANSI.

```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
