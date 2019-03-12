---
title: User-defined data type (VBA)
keywords: vblr6.chm1009052
f1_keywords:
- vblr6.chm1009052
ms.prod: office
ms.assetid: 89ef52c6-f928-d43e-ef5d-8b6b3b5a3bce
ms.date: 11/19/2018
localization_priority: Normal
---


# User-defined data type

Any [data type](../Glossary/vbe-glossary.md#data-type) that you define by using the **Type** statement.

User-defined data types can contain one or more elements of a data type, an [array](../Glossary/vbe-glossary.md#array), or a previously defined user-defined type. For example:


```vb
Type MyType 
 MyName As String ' String variable stores a name. 
 MyBirthDate As Date ' Date variable stores a birthdate. 
 MySex As Integer ' Integer variable stores sex (0 for 
End Type ' female, 1 for male). 

```

## See also

- [Data type summary](../reference/user-interface-help/data-type-summary.md)
- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
