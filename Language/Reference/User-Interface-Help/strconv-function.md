---
title: StrConv Function
keywords: vblr6.chm1011063
f1_keywords:
- vblr6.chm1011063
ms.prod: office
ms.assetid: c0b92ca2-9850-7f37-07e0-8e313646c56c
ms.date: 06/08/2017
---


# StrConv Function



Returns a  **Variant** (**String**) converted as specified.
<<<<<<< HEAD
 **Syntax**
 **StrConv** (**_string, conversion, LCID_**)
The  **StrConv** function syntax has these[named arguments](../../Glossary/vbe-glossary.md):
=======

## Syntax

**StrConv** (**_string, conversion, LCID_**)
The  **StrConv** function syntax has these[named arguments](../../Glossary/vbe-glossary.md#named-argument):
>>>>>>> master


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
|**_string_**|Required. [String expression](../../Glossary/vbe-glossary.md) to be converted.|
|**_conversion_**|Required. [Integer](../../Glossary/vbe-glossary.md). The sum of values specifying the type of conversion to perform.|
|**_LCID_**|Optional. The LocaleID, if different than the system LocaleID. (The system LocaleID is the default.)|

 **Settings**
The  **_conversion_**[argument](../../Glossary/vbe-glossary.md) settings are:
=======
|**_string_**|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) to be converted.|
|**_conversion_**|Required. [Integer](../../Glossary/vbe-glossary.md#integer-data-type). The sum of values specifying the type of conversion to perform.|
|**_LCID_**|Optional. The LocaleID, if different than the system LocaleID. (The system LocaleID is the default.)|

 **Settings**
The  **_conversion_**[argument](../../Glossary/vbe-glossary.md#argument) settings are:
>>>>>>> master


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbUpperCase**|1|Converts the string to uppercase characters.|
|**vbLowerCase**|2|Converts the string to lowercase characters.|
|**vbProperCase**|3|Converts the first letter of every word in string to uppercase.|
|**vbWide***|4*|Converts narrow (single-byte) characters in string to wide (double-byte) characters.|
|**vbNarrow***|8*|Converts wide (double-byte) characters in string to narrow (single-byte) characters.|
|**vbKatakana****|16**|Converts Hiragana characters in string to Katakana characters.|
|**vbHiragana****|32**|Converts Katakana characters in string to Hiragana characters.|
<<<<<<< HEAD
|**vbUnicode**|64|Converts the string to [Unicode](../../Glossary/vbe-glossary.md) using the default code page of the system. (Not available on the Macintosh.)|
=======
|**vbUnicode**|64|Converts the string to [Unicode](../../Glossary/vbe-glossary.md#unicode) using the default code page of the system. (Not available on the Macintosh.)|
>>>>>>> master
|**vbFromUnicode**|128|Converts the string from Unicode to the default code page of the system. (Not available on the Macintosh.)|

*Applies to East Asia locales.
**Applies to Japan only.

<<<<<<< HEAD
 **Note**  These [constants](../../Glossary/vbe-glossary.md) are specified by Visual Basic for Applications. As a result, they may be used anywhere in your code in place of the actual values. Most can be combined, for example, **vbUpperCase + vbWide**, except when they are mutually exclusive, for example, **vbUnicode + vbFromUnicode**. The constants **vbWide**, **vbNarrow**, **vbKatakana**, and **vbHiragana** cause[run-time errors](../../Glossary/vbe-glossary.md) when used in[locales](../../Glossary/vbe-glossary.md) where they do not apply.

The following are valid word separators for proper casing: [Null](../../Glossary/vbe-glossary.md) (**Chr$(** 0 **)**), horizontal tab (**Chr$(** 9 **)**), linefeed (**Chr$(** 10 **)**), vertical tab (**Chr$(** 11 **)**), form feed (**Chr$(** 12 **)**), carriage return (**Chr$(** 13 **)**), space (SBCS) (**Chr$(** 32 **)**). The actual value for a space varies by country/region for[DBCS](../../Glossary/vbe-glossary.md).
 **Remarks**
=======
 **Note**  These [constants](../../Glossary/vbe-glossary.md#constant) are specified by Visual Basic for Applications. As a result, they may be used anywhere in your code in place of the actual values. Most can be combined, for example, **vbUpperCase + vbWide**, except when they are mutually exclusive, for example, **vbUnicode + vbFromUnicode**. The constants **vbWide**, **vbNarrow**, **vbKatakana**, and **vbHiragana** cause[run-time errors](../../Glossary/vbe-glossary.md#run-time-error) when used in[locales](../../Glossary/vbe-glossary.md#locale) where they do not apply.

The following are valid word separators for proper casing: [Null](../../Glossary/vbe-glossary.md#null) (**Chr$(** 0 **)**), horizontal tab (**Chr$(** 9 **)**), linefeed (**Chr$(** 10 **)**), vertical tab (**Chr$(** 11 **)**), form feed (**Chr$(** 12 **)**), carriage return (**Chr$(** 13 **)**), space (SBCS) (**Chr$(** 32 **)**). The actual value for a space varies by country/region for[DBCS](../../Glossary/vbe-glossary.md#dbcs).

## Remarks

>>>>>>> master
When converting from a  **Byte** array in ANSI format to a string, use the **StrConv** function. When converting from such an array in Unicode format, use an assignment statement.

## Example

This example uses the  **StrConv** function to convert a Unicode string to an ANSI string.


```vb
Dim i As Long
Dim x() As Byte
x = StrConv("ABCDEFG", vbFromUnicode)    ' Convert string.
For i = 0 To UBound(x)
    Debug.Print x(i)
Next

```


