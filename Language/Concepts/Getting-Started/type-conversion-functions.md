---
<<<<<<< HEAD
title: Type Conversion Functions
=======
title: Type conversion functions
>>>>>>> master
keywords: vblr6.chm1008820
f1_keywords:
- vblr6.chm1008820
ms.prod: office
ms.assetid: fd602e34-9de2-1e8b-46fe-6a2873d6a785
<<<<<<< HEAD
ms.date: 06/08/2017
---


# Type Conversion Functions

Each function coerces an [expression](../../Glossary/vbe-glossary.md) to a specific [data type](../../Glossary/vbe-glossary.md).

 **Syntax**
=======
ms.date: 08/24/2018
---


# Type conversion functions

Each function coerces an expression to a specific data type.

## Syntax
>>>>>>> master

- **CBool(**_expression_**)**
- **CByte(**_expression_**)**
- **CCur(**_expression_**)**
- **CDate(**_expression_**)**
- **CDbl(**_expression_**)**
- **CDec(**_expression_**)**
- **CInt(**_expression_**)**
- **CLng(**_expression_**)**
- **CLngLng(**_expression_**)** (Valid on 64-bit platforms only.)
- **CLngPtr(**_expression_**)**
- **CSng(**_expression_**)**
- **CStr(**_expression_**)**
- **CVar(**_expression_**)**

<<<<<<< HEAD
The required  _expression_ [argument](../../Glossary/vbe-glossary.md) is any [string expression](../../Glossary/vbe-glossary.md) or [numeric expression](../../Glossary/vbe-glossary.md).

**Return Types**
=======
The required  _expression_ argument is any string expression or numeric expression.

### Return types
>>>>>>> master

The function name determines the return type as shown in the following:


|**Function**|**Return Type**|**Range for  _expression_ argument**|
|:-----|:-----|:-----|
<<<<<<< HEAD
|**CBool** [Boolean](../../Glossary/vbe-glossary.md)|Any valid  **string** or numeric expression.|
|**CByte** [Byte](../../Glossary/vbe-glossary.md)|0 to 255.|
|**CCur** [Currency](../../Glossary/vbe-glossary.md)|-922,337,203,685,477.5808 to 922,337,203,685,477.5807.|
|**CDate** [Date](../../Glossary/vbe-glossary.md)|Any valid [date expression](../../Glossary/vbe-glossary.md).|
|**CDbl** [Double](../../Glossary/vbe-glossary.md)|-1.79769313486231E308 to -4.94065645841247E-324 for negative values+ADs- 4.94065645841247E-324 to 1.79769313486232E308 for positive values.|
|**CDec** [Decimal](../../Glossary/vbe-glossary.md)|79,228,162,514,264,337,593,543,950,335 for zero-scaled numbers, that is, numbers with no decimal places. For numbers with 28 decimal places, the range is 7.9228162514264337593543950335. The smallest possible non-zero number is 0.0000000000000000000000000001.|
|**CInt** [Integer](../../Glossary/vbe-glossary.md)|-32,768 to 32,767+ADs- fractions are rounded.|
|**CLng** [Long](../../Glossary/vbe-glossary.md)|-2,147,483,648 to 2,147,483,647+ADs- fractions are rounded.|
|**CLngLng** [LongLong](../../reference/User-Interface-Help/longlong-data-type.md)|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807+ADs- fractions are rounded. (Valid on 64-bit platforms only.)|
|**CLngPtr** [LongPtr](../../reference/User-Interface-Help/longptr-data-type.md)|-2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems+ADs- fractions are rounded for 32-bit and 64-bit systems.|
|**CSng** [Single](../../Glossary/vbe-glossary.md)|-3.402823E38 to -1.401298E-45 for negative values+ADs- 1.401298E-45 to 3.402823E38 for positive values.|
|**CStr** [String](../../Glossary/vbe-glossary.md)|AWw-Returns for CStr](../../reference/User-Interface-Help/returns-for-cstr.md) depend on the _expression_ argument.|
|**CVar** [Variant](../../Glossary/vbe-glossary.md)|Same range as  **Double** for numerics. Same range as **String** for non-numerics.|
=======
|**CBool**|Boolean|Any valid **string** or numeric expression.|
|**CByte**|Byte|0 to 255.|
|**CCur**|Currency|-922,337,203,685,477.5808 to 922,337,203,685,477.5807.|
|**CDate**|Date|Any valid date expression.|
|**CDbl**|Double|-1.79769313486231E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values.|
|**CDec**|Decimal|79,228,162,514,264,337,593,543,950,335 for zero-scaled numbers, that is, numbers with no decimal places. For numbers with 28 decimal places, the range is 7.9228162514264337593543950335. The smallest possible non-zero number is 0.0000000000000000000000000001.|
|**CInt**|Integer|-32,768 to 32,767; fractions are rounded.|
|**CLng**|Long|-2,147,483,648 to 2,147,483,647; fractions are rounded.|
|**CLngLng**|LongLong|-9,223,372,036,854,775,808 to 9,223,372,036,854,775,807; fractions are rounded. (Valid on 64-bit platforms only.)|
|**CLngPtr**|LongPtr|-2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems; fractions are rounded for 32-bit and 64-bit systems.|
|**CSng**|Single|-3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values.|
|**CStr**|String|Returns for CStr depend on the _expression_ argument.|
|**CVar**|Variant|Same range as **Double** for numerics. Same range as **String** for non-numerics.|
>>>>>>> master

## Remarks

If the  _expression_ passed to the function is outside the range of the data type being converted to, an error occurs.

<<<<<<< HEAD
 >**Note**  Conversion functions must be used to explicitly assign  **LongLong** (including **LongPtr** on 64-bit platforms) to smaller integral types. Implicit conversions of **LongLong** to smaller integrals are not allowed.

In general, you can document your code using the data-type conversion functions to show that the result of some operation should be expressed as a particular data type rather than the default data type. For example, use  **CCur** to force currency arithmetic in cases where single-precision, double-precision, or integer arithmetic normally would occur.
You should use the data-type conversion functions instead of  **Val** to provide internationally aware conversions from one data type to another. For example, when you use **CCur**, different decimal separators, different thousand separators, and various currency options are properly recognized depending on the [locale](../../Glossary/vbe-glossary.md) setting of your computer.
When the fractional part is exactly 0.5,  **CInt** and **CLng** always round it to the nearest even number. For example, 0.5 rounds to 0, and 1.5 rounds to 2. **CInt** and **CLng** differ from the **Fix** and **Int** functions, which truncate, rather than round, the fractional part of a number. Also, **Fix** and **Int** always return a value of the same type as is passed in.
Use the  **IsDate** function to determine if _date_ can be converted to a date or time. **CDate** recognizes [date literals](../../Glossary/vbe-glossary.md) and time literals as well as some numbers that fall within the range of acceptable dates. When converting a number to a date, the whole number portion is converted to a date. Any fractional part of the number is converted to a time of day, starting at midnight.
 **CDate** recognizes date formats according to the locale setting of your system. The correct order of day, month, and year may not be determined if it is provided in a format other than one of the recognized date settings. In addition, a long date format is not recognized if it also contains the day-of-the-week string.
A  **CVDate** function is also provided for compatibility with previous versions of Visual Basic. The syntax of the **CVDate** function is identical to the **CDate** function, however, **CVDate** returns a **Variant** whose subtype is **Date** instead of an actual **Date** type. Since there is now an intrinsic **Date** type, there is no further need for **CVDate**. The same effect can be achieved by converting an expression to a **Date,** and then assigning it to a **Variant**. This technique is consistent with the conversion of all other intrinsic types to their equivalent **Variant** subtypes.

>**Note**  The  **CDec** function does not return a discrete data type+ADs- instead, it always returns a **Variant** whose value has been converted to a **Decimal** subtype.


## CBool Function Example

This example uses the  **CBool** function to convert an expression to a **Boolean**. If the expression evaluates to a nonzero value, **CBool** returns **True**A7- otherwise, it returns **False**.
=======
> [!NOTE] 
> Conversion functions must be used to explicitly assign **LongLong** (including **LongPtr** on 64-bit platforms) to smaller integral types. Implicit conversions of **LongLong** to smaller integrals are not allowed.

In general, you can document your code using the data-type conversion functions to show that the result of some operation should be expressed as a particular data type rather than the default data type. For example, use **CCur** to force currency arithmetic in cases where single-precision, double-precision, or integer arithmetic normally would occur.

You should use the data-type conversion functions instead of **Val** to provide internationally aware conversions from one data type to another. For example, when you use **CCur**, different decimal separators, different thousand separators, and various currency options are properly recognized depending on the locale setting of your computer.

When the fractional part is exactly 0.5, **CInt** and **CLng** always round it to the nearest even number. For example, 0.5 rounds to 0, and 1.5 rounds to 2. **CInt** and **CLng** differ from the [**Fix** and **Int** functions]((../../Reference/User-Interface-Help/int-fix-functions.md), which truncate, rather than round, the fractional part of a number. Also, **Fix** and **Int** always return a value of the same type as is passed in.

Use the **IsDate** function to determine if _date_ can be converted to a date or time. **CDate** recognizes date literals and time literals as well as some numbers that fall within the range of acceptable dates. When converting a number to a date, the whole number portion is converted to a date. Any fractional part of the number is converted to a time of day, starting at midnight.

**CDate** recognizes date formats according to the locale setting of your system. The correct order of day, month, and year may not be determined if it is provided in a format other than one of the recognized date settings. In addition, a long date format is not recognized if it also contains the day-of-the-week string.
 
A **CVDate** function is also provided for compatibility with previous versions of Visual Basic. The syntax of the **CVDate** function is identical to the **CDate** function; however, **CVDate** returns a **Variant** whose subtype is **Date** instead of an actual **Date** type. Since there is now an intrinsic **Date** type, there is no further need for **CVDate**. The same effect can be achieved by converting an expression to a **Date**, and then assigning it to a **Variant**. This technique is consistent with the conversion of all other intrinsic types to their equivalent **Variant** subtypes.

> [!NOTE] 
> The **CDec** function does not return a discrete data type; instead, it always returns a **Variant** whose value has been converted to a **Decimal** subtype.


## CBool function example

This example uses the **CBool** function to convert an expression to a **Boolean**. If the expression evaluates to a nonzero value, **CBool** returns **True**, otherwise, it returns **False**.
>>>>>>> master


```vb
Dim A, B, Check 
A = 5: B = 5 ' Initialize variables. 
Check = CBool(A = B) ' Check contains True. 
 
A = 0 ' Define variable. 
Check = CBool(A) ' Check contains False. 

```


<<<<<<< HEAD
## CByte Function Example

This example uses the  **CByte** function to convert an expression to a **Byte**.
=======
## CByte function example

This example uses the **CByte** function to convert an expression to a **Byte**.
>>>>>>> master


```vb
Dim MyDouble, MyByte 
MyDouble = 125.5678 ' MyDouble is a Double. 
MyByte = CByte(MyDouble) ' MyByte contains 126. 

```


<<<<<<< HEAD
## CCur Function Example

This example uses the  **CCur** function to convert an expression to a **Currency**.
=======
## CCur function example

This example uses the **CCur** function to convert an expression to a **Currency**.
>>>>>>> master


```vb
Dim MyDouble, MyCurr 
MyDouble = 543.214588 ' MyDouble is a Double. 
<<<<<<< HEAD
MyCurr = CCur(MyDouble * 2) ' Convert result of MyDouble +ACo- 2 
=======
MyCurr = CCur(MyDouble * 2) ' Convert result of MyDouble * 2 
>>>>>>> master
 ' (1086.429176) to a 
 ' Currency (1086.4292). 

```


<<<<<<< HEAD
## CDate Function Example

This example uses the  **CDate** function to convert a string to a **Date**. In general, hard-coding dates and times as strings (as shown in this example) is not recommended. Use date literals and time literals, such as +ACM-2/12/1969+ACM- and +ACM-4:45:23 PM+ACM-, instead.
=======
## CDate function example

This example uses the **CDate** function to convert a string to a **Date**. In general, hard-coding dates and times as strings (as shown in this example) is not recommended. Use date literals and time literals, such as `#2/12/1969#` and `#4:45:23 PM#`, instead.
>>>>>>> master


```vb
Dim MyDate, MyShortDate, MyTime, MyShortTime 
<<<<<<< HEAD
MyDate = February 12, 1969 ' Define date. 
MyShortDate = CDate(MyDate) ' Convert to Date data type. 
 
MyTime = 4:35:47 PM ' Define time. 
=======
MyDate = "February 12, 1969" ' Define date. 
MyShortDate = CDate(MyDate) ' Convert to Date data type. 
 
MyTime = "4:35:47 PM" ' Define time. 
>>>>>>> master
MyShortTime = CDate(MyTime) ' Convert to Date data type. 

```


<<<<<<< HEAD
## CDbl Function Example

This example uses the  **CDbl** function to convert an expression to a **Double**.
=======
## CDbl function example

This example uses the **CDbl** function to convert an expression to a **Double**.
>>>>>>> master


```vb
Dim MyCurr, MyDouble 
MyCurr = CCur(234.456784) ' MyCurr is a Currency. 
MyDouble = CDbl(MyCurr * 8.2 * 0.01) ' Convert result to a Double. 

```


<<<<<<< HEAD
## CDec Function Example

This example uses the  **CDec** function to convert a numeric value to a **Decimal**.
=======
## CDec function example

This example uses the **CDec** function to convert a numeric value to a **Decimal**.
>>>>>>> master


```vb
Dim MyDecimal, MyCurr 
MyCurr = 10000000.0587 ' MyCurr is a Currency. 
MyDecimal = CDec(MyCurr) ' MyDecimal is a Decimal. 

```


<<<<<<< HEAD
## CInt Function Example

This example uses the  **CInt** function to convert a value to an **Integer**.
=======
## CInt function example

This example uses the **CInt** function to convert a value to an **Integer**.
>>>>>>> master


```vb
Dim MyDouble, MyInt 
MyDouble = 2345.5678 ' MyDouble is a Double. 
MyInt = CInt(MyDouble) ' MyInt contains 2346. 

```


<<<<<<< HEAD
## CLng Function Example

This example uses the  **CLng** function to convert a value to a **Long**.
=======
## CLng function example

This example uses the **CLng** function to convert a value to a **Long**.
>>>>>>> master


```vb
Dim MyVal1, MyVal2, MyLong1, MyLong2 
MyVal1 = 25427.45: MyVal2 = 25427.55 ' MyVal1, MyVal2 are Doubles. 
MyLong1 = CLng(MyVal1) ' MyLong1 contains 25427. 
MyLong2 = CLng(MyVal2) ' MyLong2 contains 25428. 

```


<<<<<<< HEAD
## CSng Function Example

This example uses the  **CSng** function to convert a value to a **Single**.
=======
## CSng function example

This example uses the **CSng** function to convert a value to a **Single**.
>>>>>>> master


```vb
Dim MyDouble1, MyDouble2, MySingle1, MySingle2 
' MyDouble1, MyDouble2 are Doubles. 
MyDouble1 = 75.3421115: MyDouble2 = 75.3421555 
MySingle1 = CSng(MyDouble1) ' MySingle1 contains 75.34211. 
MySingle2 = CSng(MyDouble2) ' MySingle2 contains 75.34216. 

```


<<<<<<< HEAD
## CStr Function Example

This example uses the  **CStr** function to convert a numeric value to a **String**.
=======
## CStr function example

This example uses the **CStr** function to convert a numeric value to a **String**.
>>>>>>> master


```vb
Dim MyDouble, MyString 
MyDouble = 437.324 ' MyDouble is a Double. 
<<<<<<< HEAD
MyString = CStr(MyDouble) ' MyString contains +ACI-437.324+ACI-. 
=======
MyString = CStr(MyDouble) ' MyString contains "437.324". 
>>>>>>> master

```


<<<<<<< HEAD
## CVar Function Example

This example uses the  **CVar** function to convert an expression to a **Variant**.
=======
## CVar function example

This example uses the **CVar** function to convert an expression to a **Variant**.
>>>>>>> master


```vb
Dim MyInt, MyVar 
MyInt = 4534 ' MyInt is an Integer. 
MyVar = CVar(MyInt & 000) ' MyVar contains the string 
 ' 4534000. 

```

<<<<<<< HEAD
=======
## See also

- [Visual Basic Editor (VBE) Glossary](../../Glossary/vbe-glossary.md)
>>>>>>> master

