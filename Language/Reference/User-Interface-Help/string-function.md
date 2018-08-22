---
title: String Function
keywords: vblr6.chm1011358
f1_keywords:
- vblr6.chm1011358
ms.prod: office
ms.assetid: d6c5c054-21b9-f777-acae-ac31710ba5c5
ms.date: 06/08/2017
---


# String Function



Returns a  **Variant** (**String**) containing a repeating character string of the length specified.

## Syntax

**String** (**_number_**, **_character_**)
The  **String** function syntax has these[named arguments](../../Glossary/vbe-glossary.md#named-argument):


|**Part**|**Description**|
|:-----|:-----|
|**_number_**|Required; [Long](../../Glossary/vbe-glossary.md#Long). Length of the returned string. If  **_number_** contains[Null](../../Glossary/vbe-glossary.md#Null),  **Null** is returned.|
|**_character_**|Required; [Variant](../../Glossary/vbe-glossary.md#Variant). [Character code](../../Glossary/vbe-glossary.md#character-code) specifying the character or[string expression](../../Glossary/vbe-glossary.md#string-expression) whose first character is used to build the return string. If **_character_** contains **Null**, **Null** is returned.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
If you specify a number for  **_character_** greater than 255, **String** converts the number to a valid character code using the formula:
 **_character_** **Mod** 256

## Example

This example uses the  **String** function to return repeating character strings of the length specified.


```vb
Dim MyString
MyString = String(5, "*")    ' Returns "*****".
MyString = String(5, 42)    ' Returns "*****".
MyString = String(10, "ABC")    ' Returns "AAAAAAAAAA".


```


