---
title: Left Function
keywords: vblr6.chm1011361
f1_keywords:
- vblr6.chm1011361
ms.prod: office
ms.assetid: 2835aa57-6273-8f72-4ee8-ec19df26c5d9
ms.date: 06/08/2017
---


# Left Function



Returns a  **Variant** (**String**) containing a specified number of characters from the left side of a string.

## Syntax

**Left** (**_string_**, **_length_**)
The  **Left** function syntax has these[named arguments](../../Glossary/vbe-glossary.md#named-argument):


|**Part**|**Description**|
|:-----|:-----|
|**_string_**|Required. [String expression](../../Glossary/vbe-glossary.md#string-expression) from which the leftmost characters are returned. If **_string_** contains[Null](../../Glossary/vbe-glossary.md#null), Null is returned.|
|**_length_**|Required;  **Variant** (**Long**).[Numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) indicating how many characters to return. If 0, a zero-length string ("") is returned. If greater than or equal to the number of characters in **_string_**, the entire string is returned.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
To determine the number of characters in  **_string_**, use the **Len** function.

 **Note**  Use the  **LeftB** function with byte data contained in a string. Instead of specifying the number of characters to return, **_length_** specifies the number of bytes.


## Example

This example uses the  **Left** function to return a specified number of characters from the left side of a string.


```vb
Dim AnyString, MyStr
AnyString = "Hello World"    ' Define string.
MyStr = Left(AnyString, 1)    ' Returns "H".
MyStr = Left(AnyString, 7)    ' Returns "Hello W".
MyStr = Left(AnyString, 20)    ' Returns "Hello World".


```


