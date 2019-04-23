---
title: With statement (VBA)
keywords: vblr6.chm1009555
f1_keywords:
- vblr6.chm1009555
ms.prod: office
ms.assetid: cd548bae-ce3d-e044-7bb8-85b051a8f4a5
ms.date: 12/03/2018
localization_priority: Normal
---


# With statement

Executes a series of [statements](../../Glossary/vbe-glossary.md#statement) on a single object or a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type).

## Syntax

**With** _object_ [ _statements_ ] **End With**

<br/>

The **With** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Name of an object or a user-defined type.|
| _statements_|Optional. One or more statements to be executed on _object_.|

## Remarks

The **With** statement allows you to perform a series of statements on a specified object without requalifying the name of the object. For example, to change a number of different [properties](../../Glossary/vbe-glossary.md#property) on a single object, place the property assignment statements within the **With** control structure, referring to the object once instead of referring to it with each property assignment. 

The following example illustrates use of the **With** statement to assign values to several properties of the same object.

```vb
With MyLabel 
 .Height = 2000 
 .Width = 2000 
 .Caption = "This is MyLabel" 
End With 

```

> [!NOTE] 
> Once a **With** block is entered, _object_ can't be changed. As a result, you can't use a single **With** statement to affect a number of different objects.

You can nest **With** statements by placing one **With** block within another. However, because members of outer **With** blocks are masked within the inner **With** blocks, you must provide a fully qualified object reference in an inner **With** block to any member of an object in an outer **With** block.

> [!NOTE] 
> In general, it's recommended that you don't jump into or out of **With** blocks. If statements in a **With** block are executed, but either the **With** or **End With** statement is not executed, a temporary variable containing a reference to the object remains in memory until you exit the procedure.


## Example

This example uses the **With** statement to execute a series of statements on a single object. The object and its properties are generic names used for illustration purposes only.


```vb
With MyObject 
 .Height = 100 ' Same as MyObject.Height = 100. 
 .Caption = "Hello World" ' Same as MyObject.Caption = "Hello World". 
 With .Font 
  .Color = Red ' Same as MyObject.Font.Color = Red. 
  .Bold = True ' Same as MyObject.Font.Bold = True. 
 End With
End With
```

## See also

- [Using With statements](../../concepts/getting-started/using-with-statements.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
