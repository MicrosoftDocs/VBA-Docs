---
title: Declaring arrays (VBA)
keywords: vbcn6.chm1076697
f1_keywords:
- vbcn6.chm1076697
ms.prod: office
ms.assetid: 3efbbe80-ee1a-e660-de4b-ffb3602ac31b
ms.date: 12/21/2018
localization_priority: Priority
---


# Declaring arrays

[Arrays](../../Glossary/vbe-glossary.md#array) are declared the same way as other [variables](../../Glossary/vbe-glossary.md#variable), by using the **[Dim](../../reference/user-interface-help/dim-statement.md)**, **[Static](../../reference/user-interface-help/static-statement.md)**, **[Private](../../reference/user-interface-help/private-statement.md)**, or **[Public](../../reference/user-interface-help/public-statement.md)** statements. The difference between scalar variables (those that aren't arrays) and array variables is that you generally must specify the size of the array. An array whose size is specified is a fixed-size array. An array whose size can be changed while a program is running is a dynamic array.

Whether an array is indexed from 0 or 1 depends on the setting of the **[Option Base](../../reference/user-interface-help/option-base-statement.md)** statement. If **Option Base 1** is not specified, all array indexes begin at zero.

## Declare a fixed array

In the following line of code, a fixed-size array is declared as an **Integer** array having 11 rows and 11 columns:

```vb
Dim MyArray(10, 10) As Integer 

```

The first argument represents the rows; the second argument represents the columns.

As with any other variable declaration, unless you specify a [data type](../../reference/user-interface-help/data-type-summary.md) for the array, the data type of the elements in a declared array is **Variant**. Each numeric **Variant** element of the array uses 16 bytes. Each string **Variant** element uses 22 bytes. To write code that is as compact as possible, explicitly declare your arrays to be of a data type other than **Variant**. 

The following lines of code compare the size of several arrays.

```vb
' Integer array uses 22 bytes (11 elements * 2 bytes). 
ReDim MyIntegerArray(10) As Integer 
 
' Double-precision array uses 88 bytes (11 elements * 8 bytes). 
ReDim MyDoubleArray(10) As Double 
 
' Variant array uses at least 176 bytes (11 elements * 16 bytes). 
ReDim MyVariantArray(10) 
 
' Integer array uses 100 * 100 * 2 bytes (20,000 bytes). 
ReDim MyIntegerArray (99, 99) As Integer 
 
' Double-precision array uses 100 * 100 * 8 bytes (80,000 bytes). 
ReDim MyDoubleArray (99, 99) As Double 
 
' Variant array uses at least 160,000 bytes (100 * 100 * 16 bytes). 
ReDim MyVariantArray(99, 99) 

```

The maximum size of an array varies, based on your operating system and how much memory is available. Using an array that exceeds the amount of RAM available on your system is slower because the data must be read from and written to disk.


## Declare a dynamic array

By declaring a dynamic array, you can size the array while the code is running. Use a **Static**, **Dim**, **Private**, or **Public** statement to declare an array, leaving the parentheses empty, as shown in the following example.

```vb
Dim sngArray() As Single 

```

> [!NOTE] 
> You can use the **[ReDim](../../reference/user-interface-help/redim-statement.md)** statement to declare an array implicitly within a procedure. Be careful not to misspell the name of the array when you use the **ReDim** statement. Even if the **[Option Explicit](../../reference/user-interface-help/option-explicit-statement.md)** statement is included in the module, a second array will be created.

In a procedure within the array's [scope](../../Glossary/vbe-glossary.md#scope), use the **ReDim** statement to change the number of dimensions, to define the number of elements, and to define the upper and lower bounds for each dimension. You can use the **ReDim** statement to change the dynamic array as often as necessary. However, each time you do this, the existing values in the array are lost. Use **ReDim Preserve** to expand an array while preserving existing values in the array. 

For example, the following statement enlarges the array by 10 elements without losing the current values of the original elements.

```vb
ReDim Preserve varArray(UBound(varArray) + 10) 

```

> [!NOTE] 
> When you use the **Preserve** [keyword](../../Glossary/vbe-glossary.md#keyword) with a dynamic array, you can change only the upper bound of the last dimension, but you can't change the number of dimensions.


## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
