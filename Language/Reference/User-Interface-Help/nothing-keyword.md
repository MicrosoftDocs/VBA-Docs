---
title: Nothing keyword (VBA)
keywords: vblr6.chm1011405
f1_keywords:
- vblr6.chm1011405
ms.prod: office
ms.assetid: 9eedf4db-3aca-df26-8bc7-c3a7f7264e6b
ms.date: 05/04/2019
localization_priority: Normal
---


# Nothing keyword

The **Nothing** [keyword](../../Glossary/vbe-glossary.md#keyword) is used to disassociate an object [variable](../../Glossary/vbe-glossary.md#variable) from an actual object. Use the **[Set](set-statement.md)** statement to assign **Nothing** to an object variable. For example:

```vb
Set MyObject = Nothing 

```

Several object variables can refer to the same actual object. When **Nothing** is assigned to an object variable, that variable no longer refers to an actual object. 

When several object variables refer to the same object, memory and system resources associated with the object to which the variables refer are released only after all of them have been set to **Nothing**, either explicitly by using **Set**, or implicitly after the last object variable referencing the actual object goes out of [scope](../../Glossary/vbe-glossary.md#scope).



## See also

- [Keywords (VBA)](../keywords-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]