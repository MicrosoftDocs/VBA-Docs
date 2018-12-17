---
title: ReadAll method (Visual Basic for Applications)
keywords: vblr6.chm2182077
f1_keywords:
- vblr6.chm2182077
ms.prod: office
api_name:
- Office.ReadAll
ms.assetid: 2e461101-12ec-0472-2719-53e714632698
ms.date: 12/14/2018
---


# ReadAll method

Reads an entire **[TextStream](textstream-object.md)** file and returns the resulting string.

## Syntax

_object_.**ReadAll**

The _object_ is always the name of a **TextStream** object.

## Remarks

For large files, using the **ReadAll** method wastes memory resources. Other techniques should be used to input a file, such as reading a file line-by-line.

## See also

- [Methods (Visual Basic for Applications)](../methods-visual-basic-for-applications.md)