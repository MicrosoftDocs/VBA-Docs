---
title: ReDim Statement
keywords: vblr6.chm1008999
f1_keywords:
- vblr6.chm1008999
ms.prod: office
ms.assetid: 5044cb55-6cdc-16a7-6558-dcff7ab4b933
ms.date: 06/08/2017
---


# ReDim Statement

Used at [procedure level](../../Glossary/vbe-glossary.md#procedure-level) to reallocate storage space for dynamic array[variables](../../Glossary/vbe-glossary.md#variable).

## Syntax

**ReDim** [ **Preserve** ] _varname_**(**_subscripts_**)** [ **As**_type_ ] [ **,**_varname_**(**_subscripts_**)** [ **As**_type_ ]] **. . .**

The  **ReDim** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Preserve**|Optional. [Keyword](../../Glossary/vbe-glossary.md#Keyword) used to preserve the data in an existing[array](../../Glossary/vbe-glossary.md#array) when you change the size of the last dimension.|
| _varname_|Required. Name of the variable; follows standard variable naming conventions.|
| _subscripts_|Required. Dimensions of an array variable; up to 60 multiple dimensions may be declared. The  _subscripts_[argument](../../Glossary/vbe-glossary.md#argument) uses the following syntax: [ _lower_**To** ] _upper_ [ **,** [ _lower_**To** ] _upper_ ] **. . .** When not explicitly stated in _lower_, the lower bound of an array is controlled by the **Option** **Base** statement. The lower bound is zero if no **Option** **Base** statement is present.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#Data-type) of the variable; may be[Byte](../../Glossary/vbe-glossary.md#Byte), [Boolean](../../Glossary/vbe-glossary.md#Boolean), [Integer](../../Glossary/vbe-glossary.md#Integer), [Long](../../Glossary/vbe-glossary.md#Long), [Currency](../../Glossary/vbe-glossary.md#Currency), [Single](../../Glossary/vbe-glossary.md#Single), [Double](../../Glossary/vbe-glossary.md#Double), [Decimal](../../Glossary/vbe-glossary.md#Decimal) (not currently supported),[Date](../../Glossary/vbe-glossary.md#Date), [String](../../Glossary/vbe-glossary.md#String) (for variable-length strings), **String** * _length_ (for fixed-length strings),[Object](../../Glossary/vbe-glossary.md#Object), [Variant](../../Glossary/vbe-glossary.md#Variant), a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), or an [object type](../../Glossary/vbe-glossary.md#object-type). Use a separate  **As**_type_ clause for each variable being defined. For a **Variant** containing an array, _type_ describes the type of each element of the array, but doesn't change the **Variant** to some other type.|

## Remarks

The  **ReDim**[statement](../../Glossary/vbe-glossary.md#statement) is used to size or resize a dynamic array that has already been formally declared using a **Private**, **Public**, or **Dim** statement with empty parentheses (without dimension subscripts).
You can use the  **ReDim** statement repeatedly to change the number of elements and dimensions in an array. However, you can't declare an array of one data type and later use **ReDim** to change the array to another data type, unless the array is contained in a **Variant**. If the array is contained in a **Variant**, the type of the elements can be changed using an **As**_type_ clause, unless you're using the **Preserve** keyword, in which case, no changes of data type are permitted.
If you use the  **Preserve** keyword, you can resize only the last array dimension and you can't change the number of dimensions at all. For example, if your array has only one dimension, you can resize that dimension because it is the last and only dimension. However, if your array has two or more dimensions, you can change the size of only the last dimension and still preserve the contents of the array. The following example shows how you can increase the size of the last dimension of a dynamic array without erasing any existing data contained in the array.



```vb
ReDim X(10, 10, 10) 
. . . 
ReDim Preserve X(10, 10, 15) 

```

Similarly, when you use  **Preserve**, you can change the size of the array only by changing the upper bound; changing the lower bound causes an error.
If you make an array smaller than it was, data in the eliminated elements will be lost. If you pass an array to a procedure by reference, you can't redimension the array within the procedure.
When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros.  **Variant** variables are initialized to[Empty](../../Glossary/vbe-glossary.md#Empty). Each element of a user-defined type variable is initialized as if it were a separate variable. A variable that refers to an object must be assigned an existing object using the  **Set** statement before it can be used. Until it is assigned an object, the declared[object variable](../../Glossary/vbe-glossary.md#object-variable) has the special value **Nothing**, which indicates that it doesn't refer to any particular instance of an object.
The  **ReDim** statement acts as a declarative statement if the variable it declares doesn't exist at[module level](../../Glossary/vbe-glossary.md#module-level) or[procedure level](../../Glossary/vbe-glossary.md#procedure-level). If another variable with the same name is created later, even in a wider [scope](../../Glossary/vbe-glossary.md#scope),  **ReDim** will refer to the later variable and won't necessarily cause a compilation error, even if **Option Explicit** is in effect. To avoid such conflicts, **ReDim** should not be used as a declarative statement, but simply for redimensioning arrays.

 **Note**  To resize an array contained in a  **Variant**, you must explicitly declare the **Variant** variable before attempting to resize its array.


## Example

This example uses the  **ReDim** statement to allocate and reallocate storage space for dynamic-array variables. It assumes the **Option Base** is **1**.


```vb
Dim MyArray() As Integer ' Declare dynamic array. 
Redim MyArray(5) ' Allocate 5 elements. 
For I = 1 To 5 ' Loop 5 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

The next statement resizes the array and erases the elements.




```vb
Redim MyArray(10) ' Resize to 10 elements. 
For I = 1 To 10 ' Loop 10 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

The following statement resizes the array but does not erase elements.




```vb
Redim Preserve MyArray(15) ' Resize to 15 elements. 

```


