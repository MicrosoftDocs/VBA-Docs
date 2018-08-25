---
title: Right Function
keywords: vblr6.chm1011365
f1_keywords:
- vblr6.chm1011365
ms.prod: office
ms.assetid: efa00f0a-8d7d-df81-f889-16de010c2f53
ms.date: 06/08/2017
---


# Right Function



Returns a  **Variant** (**String**) containing a specified number of characters from the right side of a string.

## Syntax

**Right** (**_string_**, **_length_**)
The  **Right** function syntax has these[named arguments](../../Glossary/vbe-glossary.md#named-argument):


|**Part**|**Description**|
|:-----|:-----|
|**_string_**|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) from which the rightmost characters are returned. If **_string_** contains[Null](../../Glossary/vbe-glossary.md#null),  **Null** is returned.|
|**_length_**|Required;  **Variant** (**Long**).[Numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) indicating how many characters to return. If 0, a zero-length string ("") is returned. If greater than or equal to the number of characters in **_string_**, the entire string is returned.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
To determine the number of characters in  **_string_**, use the **Len** function.

 **Note**  Use the  **RightB** function with byte data contained in a string. Instead of specifying the number of characters to return, **_length_** specifies the number of bytes.


## Example

This example uses the  **Right** function to return a specified number of characters from the right side of a string.


```vb
Dim AnyString, MyStr
AnyString = "Hello World"    ' Define string.
MyStr = Right(AnyString, 1)    ' Returns "d".
MyStr = Right(AnyString, 6)    ' Returns " World".
MyStr = Right(AnyString, 20)    ' Returns "Hello World".


```


