---
title: Data type summary
keywords: vblr6.chm1008885
f1_keywords:
- vblr6.chm1008885
ms.prod: office
ms.assetid: 24723bdf-8454-f661-7914-d731e74d2e7b
ms.date: 11/19/2018 
localization_priority: Priority
---


# Data type summary

A data type is the characteristic of a variable that determines what kind of data it can hold. Data types include those in the following table as well as user-defined types and specific types of objects.

## Set intrinsic data types

The following table shows the supported [data types](../../Glossary/vbe-glossary.md#data-type), including storage sizes and ranges.

|Data type|Storage size|Range|
|:--------|:-----------|:----|
|**[Boolean](boolean-data-type.md)**|2 bytes|**True** or **False**|
|**[Byte](byte-data-type.md)**|1 byte|0 to 255|
|**Collection**|Unknown|Unknown|
|**[Currency](currency-data-type.md)** (scaled integer)|8 bytes|-922,337,203,685,477.5808 to 922,337,203,685,477.5807|
|**[Date](date-data-type.md)**|8 bytes|January 1, 100, to December 31, 9999|
|**[Decimal](decimal-data-type.md)**|14 bytes|+/-79,228,162,514,264,337,593,543,950,335 with no decimal point<br/><br/>+/-7.9228162514264337593543950335 with 28 places to the right of the decimal<br/><br/>Smallest non-zero number is+/-0.0000000000000000000000000001|
|**Dictionary**|Unknown|Unknown|
|**[Double](double-data-type.md)** (double-precision floating-point)|8 bytes|-1.79769313486231E308 to -4.94065645841247E-324 for negative values<br/><br/>4.94065645841247E-324 to 1.79769313486232E308 for positive values|
|**[Integer](integer-data-type.md)**|2 bytes|-32,768 to 32,767|
|**[Long](long-data-type.md)** (Long integer)|4 bytes|-2,147,483,648 to 2,147,483,647|
|**[LongLong](longlong-data-type.md)** (LongLong integer)|8 bytes|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807<br/><br/>Valid on 64-bit platforms only.|
|**[LongPtr](longptr-data-type.md)** (Long integer on 32-bit systems, LongLong integer on 64-bit systems)|4 bytes on 32-bit systems<br/><br/>8 bytes on 64-bit systems|-2,147,483,648 to 2,147,483,647 on 32-bit systems<br/><br/>-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems|
|**[Object](object-data-type.md)**|4 bytes|Any **Object** reference|
|**[Single](single-data-type.md)** (single-precision floating-point)|4 bytes|-3.402823E38 to -1.401298E-45 for negative values<br/><br/>1.401298E-45 to 3.402823E38 for positive values|
|**[String](string-data-type.md)** (variable-length)|10 bytes + string length|0 to approximately 2 billion|
|**String** (fixed-length)|Length of string|1 to approximately 65,400|
|**[Variant](variant-data-type.md)** (with numbers)|16 bytes|Any numeric value up to the range of a **Double**|
|**Variant** (with characters)|22 bytes + string length (24 bytes on 64-bit systems)|Same range as for variable-length **String**|
|**[User-defined](../../How-to/user-defined-data-type.md)** (using **Type**) |Number required by elements|The range of each element is the same as the range of its data type.|


<br/>

A **Variant** containing an array requires 12 bytes more than the array alone.

> [!NOTE] 
> [Arrays](../../Glossary/vbe-glossary.md#array) of any data type require 20 bytes of memory plus 4 bytes for each array dimension plus the number of bytes occupied by the data itself. The memory occupied by the data can be calculated by multiplying the number of data elements by the size of each element.
> 
> For example, the data in a single-dimension array consisting of 4 **Integer** data elements of 2 bytes each occupies 8 bytes. The 8 bytes required for the data plus the 24 bytes of overhead brings the total memory requirement for the array to 32 bytes. On 64-bit platforms, SAFEARRAY's take up 24-bits (plus 4 bytes per Dim statement). The pvData member is an 8-byte pointer and it must be aligned on 8 byte boundaries.

> [!NOTE] 
> [LongPtr](longptr-data-type.md) is not a true data type because it transforms to a [Long](long-data-type.md) in 32-bit environments, or a [LongLong](longlong-data-type.md) in 64-bit environments. **LongPtr** should be used to represent pointer and handle values in [Declare statements](declare-statement.md) and enables writing portable code that can run in both 32-bit and 64-bit environments.

> [!NOTE] 
> Use the **StrConv** function to convert one type of string data to another.


## Convert between data types

See [Type conversion functions](../../concepts/getting-started/type-conversion-functions.md) for examples of how to use the following functions to coerce an expression to a specific data type: **CBool**, **CByte**, **CCur**, **CDate**, **CDbl**, **CDec**, **CInt**, **CLng**, **CLngLng**, **CLngPtr**, **CSng**, **[CStr](#return-values-for-cstr)**, and **CVar**.

For the following, see the respective function pages: **[CVErr](cverr-function.md)**, **[Fix](int-fix-functions.md)**, and **[Int](int-fix-functions.md)**.

> [!NOTE] 
> **CLngLng** is valid on 64-bit platforms only.

## Verify data types

To verify data types, see the following functions: 

- [IsArray](isarray-function.md)
- [IsDate](isdate-function.md)
- [IsEmpty](isempty-function.md)
- [IsError](iserror-function.md)
- [IsMissing](ismissing-function.md)
- [IsNull](isnull-function.md)
- [IsNumeric](isnumeric-function.md)
- [IsObject](isobject-function.md)

## Return values for CStr

|If _expression_ is|CStr returns|
|:-----------------|:-----------|
|**Boolean**|A string containing **True** or **False**.|
|**Date**|A string containing a date in the short date format of your system.|
|[Empty](../../Glossary/vbe-glossary.md#empty)|A zero-length string ("").|
|**Error**|A string containing the word **Error** followed by the [error number](../../Glossary/vbe-glossary.md#error-number).|
|[Null](../../Glossary/vbe-glossary.md#null)|A [run-time error](../../Glossary/vbe-glossary.md#run-time-error).|
|Other numeric|A string containing the number.|

## See also

- [VarType constants](../../concepts/getting-started/vartype-constants.md)
- [Keywords by task](keywords-by-task.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
