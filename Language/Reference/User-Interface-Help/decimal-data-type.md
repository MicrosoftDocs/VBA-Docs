---
title: Decimal data type
keywords: vblr6.chm1099868
f1_keywords:
- vblr6.chm1099868
ms.prod: office
ms.assetid: 5f70e06b-61da-e0be-9f96-7dd84f377c74
ms.date: 06/07/2019
localization_priority: Normal
---


# Decimal data type

[Decimal variables](../../Glossary/vbe-glossary.md#decimal-data-type) are stored as 96-bit (12-byte) unsigned integers, together with a scaling factor (used to indicate either a whole number power of 10 to scale the integer down by, or that there should be no scaling) and a value indicating whether the decimal number is positive or negative. 

The scaling factor is the number of digits to store to the right of the decimal point, and ranges from 0 to 28.

- With a scale of 0 (no decimal places), the largest possible value is +/-79,228,162,514,264,337,593,543,950,335. 

- With a scale of 28 decimal places, the largest value is +/-7.9228162514264337593543950335 and the smallest, non-zero value is +/-0.0000000000000000000000000001.


> [!NOTE] 
> At this time, the **Decimal** data type can only be used within a [Variant](../../Glossary/vbe-glossary.md#variant-data-type); that is, you cannot declare a variable to be of type **Decimal**. You can, however, create a **Variant** whose subtype is **Decimal** by using the **CDec** function.

## See also

- [2.2.26 Decimal](https://docs.microsoft.com/openspecs/windows_protocols/ms-oaut/b5493025-e447-4109-93a8-ac29c48d018d)
- [Data type summary](data-type-summary.md)
- [Data types keyword summary](data-types-keyword-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
