---
title: Boolean Data Type
keywords: vblr6.chm1009278
f1_keywords:
- vblr6.chm1009278
ms.prod: office
ms.assetid: 4c0e4d2a-5cc3-c763-cb87-7bd5c2eb82b3
ms.date: 06/08/2017
---


# Boolean Data Type

[Boolean variables](../../Glossary/vbe-glossary.md#Boolean-variable) are stored as 16-bit (2-byte) numbers, but they can only be **True** or **False**. **Boolean** variables display as either `True` or `False` (when **Print** is used) or `#TRUE#` or or `False` (when **Print** is used) or `#TRUE#` or `#FALSE#` (when **Write #** is used). Use the [keywords](../../Glossary/vbe-glossary.md#keyword) **True** and **False** to assign one of the two states to **Boolean** variables.

When other [numeric types](../../Glossary/vbe-glossary.md#numeric-type) are converted to **Boolean** values, 0 becomes **False** and all other values become **True**. When **Boolean** values are converted to other [data types](../../Glossary/vbe-glossary.md#data-type), **False** becomes 0 and **True** becomes -1.

