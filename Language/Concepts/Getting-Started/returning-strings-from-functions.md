---
title: Returning Strings from Functions
keywords: vbcn6.chm1009791
f1_keywords:
- vbcn6.chm1009791
ms.prod: office
ms.assetid: 7d344b4f-e262-7f3c-71e0-7e4a884db54e
ms.date: 06/08/2017
---


# Returning Strings from Functions

<<<<<<< HEAD
Some functions have two versions: one that returns a [Variant data type](../../Glossary/vbe-glossary.md) and one that returns a [String data type](../../Glossary/vbe-glossary.md). The  **Variant** versions are more convenient because variants handle conversions between different types of data automatically. They also allow [Null](../../Glossary/vbe-glossary.md) to be propagated through an [expression](../../Glossary/vbe-glossary.md). The  **String** versions are more efficient because they use less memory.
=======
Some functions have two versions: one that returns a [Variant data type](../../Glossary/vbe-glossary.md#variant-data-type) and one that returns a [String data type](../../Glossary/vbe-glossary.md#string-data-type). The  **Variant** versions are more convenient because variants handle conversions between different types of data automatically. They also allow [Null](../../Glossary/vbe-glossary.md#null) to be propagated through an [expression](../../Glossary/vbe-glossary.md#expression). The  **String** versions are more efficient because they use less memory.
>>>>>>> master

Consider using the  **String** version when:




<<<<<<< HEAD
- Your program is very large and uses many [variables](../../Glossary/vbe-glossary.md).
=======
- Your program is very large and uses many [variables](../../Glossary/vbe-glossary.md#variable).
>>>>>>> master
    
- You write data directly to random-access files.
    

The following functions return values in a  **String** variable when you append a dollar sign (**$**) to the function name. These functions have the same usage and syntax as their **Variant** equivalents without the dollar sign.

|**Function**|||
|:-----|:-----|:-----|
<<<<<<< HEAD
|[Chr$](../../Glossary/vbe-glossary.md)|[ChrB$](../../Glossary/vbe-glossary.md)|*[Command$](../../Glossary/vbe-glossary.md)|
|[CurDir$](../../Glossary/vbe-glossary.md)|[Date$](../../Glossary/vbe-glossary.md)|[Dir$](../../Glossary/vbe-glossary.md)|
|[Error$](../../Glossary/vbe-glossary.md)|[Format$](../../Glossary/vbe-glossary.md)|[Hex$](../../Glossary/vbe-glossary.md)|
|[Input$](../../Glossary/vbe-glossary.md)|[InputB$](../../Glossary/vbe-glossary.md)|[LCase$](../../Glossary/vbe-glossary.md)|
|[Left$](../../Glossary/vbe-glossary.md)|[LeftB$](../../Glossary/vbe-glossary.md)|[LTrim$](../../Glossary/vbe-glossary.md)|
|[Mid$](../../Glossary/vbe-glossary.md)|[MidB$](../../Glossary/vbe-glossary.md)|[Oct$](../../Glossary/vbe-glossary.md)|
|[Right$](../../Glossary/vbe-glossary.md)|[RightB$](../../Glossary/vbe-glossary.md)|[RTrim$](../../Glossary/vbe-glossary.md)|
|[Space$](../../Glossary/vbe-glossary.md)|[Str$](../../Glossary/vbe-glossary.md)|[String$](../../Glossary/vbe-glossary.md)|
|[Time$](../../Glossary/vbe-glossary.md)|[Trim$](../../Glossary/vbe-glossary.md)|[UCase$](../../Glossary/vbe-glossary.md)|
=======
|[Chr$](../../Reference/User-Interface-Help/chr-function.md)|[ChrB$](../../Reference/User-Interface-Help/chr-function.md)|*[Command$](../../Reference/User-Interface-Help/command-function.md)|
|[CurDir$](../../Reference/User-Interface-Help/curdir-function.md)|[Date$](../../Reference/User-Interface-Help/date-function.md)|[Dir$](../../Reference/User-Interface-Help/dir-function.md)|
|[Error$](../../Reference/User-Interface-Help/error-function.md)|[Format$](../../Reference/User-Interface-Help/format-function-visual-basic-for-applications.md)|[Hex$](../../Reference/User-Interface-Help/hex-function.md)|
|[Input$](../../Reference/User-Interface-Help/input-function.md)|[InputB$](../../Reference/User-Interface-Help/input-function.md)|[LCase$](../../Reference/User-Interface-Help/lcase-function.md)|
|[Left$](../../Reference/User-Interface-Help/left-function.md)|[LeftB$](../../Reference/User-Interface-Help/left-function.md)|[LTrim$](../../Reference/User-Interface-Help/ltrim-rtrim-and-trim-functions.md)|
|[Mid$](../../Reference/User-Interface-Help/mid-function.md)|[MidB$](../../Reference/User-Interface-Help/mid-function.md)|[Oct$](../../Reference/User-Interface-Help/oct-function.md)|
|[Right$](../../Reference/User-Interface-Help/right-function.md)|[RightB$](../../Reference/User-Interface-Help/right-function.md)|[RTrim$](../../Reference/User-Interface-Help/ltrim-rtrim-and-trim-functions.md)|
|[Space$](../../Reference/User-Interface-Help/space-function.md)|[Str$](../../Reference/User-Interface-Help/str-function.md)|[String$](../../Reference/User-Interface-Help/string-function.md)|
|[Time$](../../Reference/User-Interface-Help/time-function.md)|[Trim$](../../Reference/User-Interface-Help/ltrim-rtrim-and-trim-functions.md)|[UCase$](../../Reference/User-Interface-Help/ucase-function.md)|
>>>>>>> master


* May not be available in all applications.

