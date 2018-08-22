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

Some functions have two versions: one that returns a [Variant data type](../../Glossary/vbe-glossary.md) and one that returns a [String data type](../../Glossary/vbe-glossary.md). The  **Variant** versions are more convenient because variants handle conversions between different types of data automatically. They also allow [Null](../../Glossary/vbe-glossary.md) to be propagated through an [expression](../../Glossary/vbe-glossary.md). The  **String** versions are more efficient because they use less memory.

Consider using the  **String** version when:




- Your program is very large and uses many [variables](../../Glossary/vbe-glossary.md).
    
- You write data directly to random-access files.
    

The following functions return values in a  **String** variable when you append a dollar sign (**$**) to the function name. These functions have the same usage and syntax as their **Variant** equivalents without the dollar sign.

|**Function**|||
|:-----|:-----|:-----|
|[Chr$](../../Glossary/vbe-glossary.md)|[ChrB$](../../Glossary/vbe-glossary.md)|*[Command$](../../Glossary/vbe-glossary.md)|
|[CurDir$](../../Glossary/vbe-glossary.md)|[Date$](../../Glossary/vbe-glossary.md)|[Dir$](../../Glossary/vbe-glossary.md)|
|[Error$](../../Glossary/vbe-glossary.md)|[Format$](../../Glossary/vbe-glossary.md)|[Hex$](../../Glossary/vbe-glossary.md)|
|[Input$](../../Glossary/vbe-glossary.md)|[InputB$](../../Glossary/vbe-glossary.md)|[LCase$](../../Glossary/vbe-glossary.md)|
|[Left$](../../Glossary/vbe-glossary.md)|[LeftB$](../../Glossary/vbe-glossary.md)|[LTrim$](../../Glossary/vbe-glossary.md)|
|[Mid$](../../Glossary/vbe-glossary.md)|[MidB$](../../Glossary/vbe-glossary.md)|[Oct$](../../Glossary/vbe-glossary.md)|
|[Right$](../../Glossary/vbe-glossary.md)|[RightB$](../../Glossary/vbe-glossary.md)|[RTrim$](../../Glossary/vbe-glossary.md)|
|[Space$](../../Glossary/vbe-glossary.md)|[Str$](../../Glossary/vbe-glossary.md)|[String$](../../Glossary/vbe-glossary.md)|
|[Time$](../../Glossary/vbe-glossary.md)|[Trim$](../../Glossary/vbe-glossary.md)|[UCase$](../../Glossary/vbe-glossary.md)|


* May not be available in all applications.

