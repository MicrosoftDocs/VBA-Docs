---
title: Currency data type
keywords: vblr6.chm1008882
f1_keywords:
- vblr6.chm1008882
ms.prod: office
ms.assetid: 4eae26dd-66c3-0181-78f9-6b59d45c19a1
ms.date: 11/19/2018
localization_priority: Normal
---


# Currency data type

[Currency variables](../../Glossary/vbe-glossary.md#currency-data-type) are stored as 64-bit (8-byte) numbers in an integer format, scaled by 10,000 to give a fixed-point number with 15 digits to the left of the decimal point and 4 digits to the right. 

This representation provides a range of -922,337,203,685,477.5808 to 922,337,203,685,477.5807. 

The [type-declaration character](../../Glossary/vbe-glossary.md#type-declaration-character) for **Currency** is the at (**@**) sign.

The **Currency** [data type](../../Glossary/vbe-glossary.md#data-type) is useful for calculations involving money and for fixed-point calculations in which accuracy is particularly important.

## See also

- [Data type summary](data-type-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
