---
title: IsObject function (Visual Basic for Applications)
keywords: vblr6.chm1008825
f1_keywords:
- vblr6.chm1008825
ms.prod: office
ms.assetid: 24fee32f-52ed-48b3-a52e-9a66b0e62723
ms.date: 12/13/2018
localization_priority: Normal
---


# IsObject function

Returns a **Boolean** value indicating whether an [identifier](../../Glossary/vbe-glossary.md#identifier) represents an object [variable](../../Glossary/vbe-glossary.md#variable).

## Syntax

**IsObject**(_identifier_)

The required _identifier_ [argument](../../Glossary/vbe-glossary.md#argument) is a variable name.

## Remarks

**IsObject** is useful only in determining whether a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) is of **VarType  vbObject**. This could occur if the **Variant** actually references (or once referenced) an object, or if it contains **[Nothing](nothing-keyword.md).**

**IsObject** returns **True** if _identifier_ is a variable declared with [Object](../../Glossary/vbe-glossary.md#object) type or any valid [class](../../Glossary/vbe-glossary.md#class) type, or if _identifier_ is a **Variant** of **VarType vbObject**, or a user-defined object; otherwise, it returns **False**. 

**IsObject** returns **True** even if the variable has been set to **Nothing**. Use error trapping to be sure that an object reference is valid.

> [!NOTE] 
> This function is useful in error handling sections of the code where you are not sure whether an object was instantiated before the error occurred, and for example, you want to close it.

## Example

This example uses the **IsObject** function to determine if an identifier represents an object variable. _MyObject_ and _YourObject_ are object variables of the same type. They are generic names used for illustration purposes only.


```vb
Dim MyInt As Integer              ' Declare variables.
Dim YourObject, MyCheck           ' Note: Default variable type is Variant.
Dim MyObject As Object
Set YourObject = MyObject         ' Assign an object reference.
MyCheck = IsObject(YourObject)    ' Returns True.
MyCheck = IsObject(MyInt)         ' Returns False.
MyCheck = IsObject(Nothing)       ' Returns True.
MyCheck = IsObject(Empty)         ' Returns False.
MyCheck = IsObject(Null)          ' Returns False.
```

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
