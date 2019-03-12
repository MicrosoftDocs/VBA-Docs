---
title: LongPtr data type
keywords: vblr6.chm1009053
f1_keywords:
- vblr6.chm1009053
ms.prod: office
ms.assetid: 10ee4c07-b686-5b86-5cea-250a9218e7ba
ms.date: 11/19/2018
localization_priority: Normal
---


# LongPtr data type

**LongPtr** ([Long](long-data-type.md) integer on 32-bit systems, [LongLong](longlong-data-type.md) integer on 64-bit systems) variables are stored as:

- Signed 32-bit (4-byte) numbers ranging in value from -2,147,483,648 to 2,147,483,647 on 32-bit systems

- Signed 64-bit (8-byte) numbers ranging in value from -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems

> [!NOTE] 
> [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. Using **LongPtr** enables writing portable code that can run in both 32-bit and 64-bit environments. Use **LongPtr** for pointers and handles.

## See also

- [Data type summary](data-type-summary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
