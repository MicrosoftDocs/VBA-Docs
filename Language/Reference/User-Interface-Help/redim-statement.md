---
title: ReDim statement (VBA)
keywords: vblr6.chm1008999
f1_keywords:
- vblr6.chm1008999
ms.prod: office
ms.assetid: 5044cb55-6cdc-16a7-6558-dcff7ab4b933
ms.date: 06/27/2019
localization_priority: Normal
---


# ReDim statement

Used at the [procedure level](../../Glossary/vbe-glossary.md#procedure-level) to reallocate storage space for dynamic array [variables](../../Glossary/vbe-glossary.md#variable).

## Syntax

**ReDim** [ **Preserve** ] _varname_ ( _subscripts_ ) [ **As** _type_ ], [ _varname_ ( _subscripts_ ) [ **As** _type_ ]] **. . .**

<br/>

The **ReDim** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**Preserve**|Optional. [Keyword](../../Glossary/vbe-glossary.md#keyword) used to preserve the data in an existing [array](../../Glossary/vbe-glossary.md#array) when you change the size of the last dimension.|
| _varname_|Required. Name of the variable; follows standard variable naming conventions.|
| _subscripts_|Required. Dimensions of an array variable; up to 60 multiple dimensions may be declared. The  _subscripts_ [argument](../../Glossary/vbe-glossary.md#argument) uses the following syntax:<br/><br/>[ _lower_**To** ] _upper_ [ , [ _lower_**To** ] _upper_ ] **. . .**<br/><br/>When not explicitly stated in _lower_, the lower bound of an array is controlled by the **[Option Base](option-base-statement.md)** statement. The lower bound is zero if no **Option Base** statement is present.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#data-type) of the variable; may be [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported), [Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (for variable-length strings), **String** _length_ (for fixed-length strings), [Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), or an [object type](../../Glossary/vbe-glossary.md#object-type).<br/><br/>Use a separate **As** _type_ clause for each variable being defined. For a **Variant** containing an array, _type_ describes the type of each element of the array, but doesn't change the **Variant** to some other type.|

## Remarks

The **ReDim** statement is used to size or resize a dynamic array that has already been formally declared by using a **[Private](private-statement.md)**, **[Public](public-statement.md)**, or **[Dim](dim-statement.md)** statement with empty parentheses (without dimension subscripts).

You can use the **ReDim** statement repeatedly to change the number of elements and dimensions in an array. However, you can't declare an array of one data type and later use **ReDim** to change the array to another data type, unless the array is contained in a **Variant**. If the array is contained in a **Variant**, the type of the elements can be changed by using an **As** _type_ clause, unless you are using the **Preserve** keyword, in which case, no changes of data type are permitted.

If you use the **Preserve** keyword, you can resize only the last array dimension and you can't change the number of dimensions at all. For example, if your array has only one dimension, you can resize that dimension because it is the last and only dimension. However, if your array has two or more dimensions, you can change the size of only the last dimension and still preserve the contents of the array. 

The following example shows how you can increase the size of the last dimension of a dynamic array without erasing any existing data contained in the array.

```vb
ReDim X(10, 10, 10) 
. . . 
ReDim Preserve X(10, 10, 15) 

```

Similarly, when you use **Preserve**, you can change the size of the array only by changing the upper bound; changing the lower bound causes an error.

If you make an array smaller than it was, data in the eliminated elements will be lost. 

When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros. **Variant** variables are initialized to [Empty](../../Glossary/vbe-glossary.md#empty). Each element of a user-defined type variable is initialized as if it were a separate variable. 

A variable that refers to an object must be assigned an existing object by using the **[Set](set-statement.md)** statement before it can be used. Until it is assigned an object, the declared [object variable](../../Glossary/vbe-glossary.md#object-variable) has the special value **[Nothing](nothing-keyword.md)**, which indicates that it doesn't refer to any particular instance of an object.

The **ReDim** statement acts as a declarative statement if the variable it declares doesn't exist at the [module level](../../Glossary/vbe-glossary.md#module-level) or [procedure level](../../Glossary/vbe-glossary.md#procedure-level). If another variable with the same name is created later, even in a wider [scope](../../Glossary/vbe-glossary.md#scope), **ReDim** will refer to the later variable and won't necessarily cause a compilation error, even if **Option Explicit** is in effect. To avoid such conflicts, **ReDim** should not be used as a declarative statement, but simply for redimensioning arrays.

> [!NOTE]
> To resize an array contained in a **Variant**, you must explicitly declare the **Variant** variable before attempting to resize its array.


## Example

This example uses the **ReDim** statement to allocate and reallocate storage space for dynamic-array variables. It assumes the **Option Base** is **1**.

```vb
Dim MyArray() As Integer ' Declare dynamic array. 
Redim MyArray(5) ' Allocate 5 elements. 
For I = 1 To 5 ' Loop 5 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

<br/>

The next statement resizes the array and erases the elements.

```vb
Redim MyArray(10) ' Resize to 10 elements. 
For I = 1 To 10 ' Loop 10 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

<br/>

The following statement resizes the array but does not erase elements.

```vb
Redim Preserve MyArray(15) ' Resize to 15 elements. 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
