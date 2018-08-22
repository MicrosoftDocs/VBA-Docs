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

Some functions have two versions: one that returns a [Variant data type](../../Glossary/vbe-glossary.md#Variant-data-type) and one that returns a [String data type](../../Glossary/vbe-glossary.md#String-data-type). The  **Variant** versions are more convenient because variants handle conversions between different types of data automatically. They also allow [Null](../../Glossary/vbe-glossary.md#Null) to be propagated through an [expression](../../Glossary/vbe-glossary.md#expression). The  **String** versions are more efficient because they use less memory.

Consider using the  **String** version when:




- Your program is very large and uses many [variables](../../Glossary/vbe-glossary.md#variable).
    
- You write data directly to random-access files.
    

The following functions return values in a  **String** variable when you append a dollar sign (**$**) to the function name. These functions have the same usage and syntax as their **Variant** equivalents without the dollar sign.

|**Function**|||
|:-----|:-----|:-----|
|[Chr$](../../Glossary/vbe-glossary.md#Chr$)|[ChrB$](../../Glossary/vbe-glossary.md#ChrB$)|*[Command$](../../Glossary/vbe-glossary.md#Command$)|
|[CurDir$](../../Glossary/vbe-glossary.md#CurDir$)|[Date$](../../Glossary/vbe-glossary.md#Date$)|[Dir$](../../Glossary/vbe-glossary.md#Dir$)|
|[Error$](../../Glossary/vbe-glossary.md#Error$)|[Format$](../../Glossary/vbe-glossary.md#Format$)|[Hex$](../../Glossary/vbe-glossary.md#Hex$)|
|[Input$](../../Glossary/vbe-glossary.md#Input$)|[InputB$](../../Glossary/vbe-glossary.md#InputB$)|[LCase$](../../Glossary/vbe-glossary.md#LCase$)|
|[Left$](../../Glossary/vbe-glossary.md#Left$)|[LeftB$](../../Glossary/vbe-glossary.md#LeftB$)|[LTrim$](../../Glossary/vbe-glossary.md#LTrim$)|
|[Mid$](../../Glossary/vbe-glossary.md#Mid$)|[MidB$](../../Glossary/vbe-glossary.md#MidB$)|[Oct$](../../Glossary/vbe-glossary.md#Oct$)|
|[Right$](../../Glossary/vbe-glossary.md#Right$)|[RightB$](../../Glossary/vbe-glossary.md#RightB$)|[RTrim$](../../Glossary/vbe-glossary.md#RTrim$)|
|[Space$](../../Glossary/vbe-glossary.md#Space$)|[Str$](../../Glossary/vbe-glossary.md#Str$)|[String$](../../Glossary/vbe-glossary.md#String$)|
|[Time$](../../Glossary/vbe-glossary.md#Time$)|[Trim$](../../Glossary/vbe-glossary.md#Trim$)|[UCase$](../../Glossary/vbe-glossary.md#UCase$)|


* May not be available in all applications.

