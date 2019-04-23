---
title: Private statement (VBA)
keywords: vblr6.chm1010962
f1_keywords:
- vblr6.chm1010962
ms.prod: office
ms.assetid: f578a258-aac1-3dc5-ab1d-e74baaaf7244
ms.date: 12/03/2018
localization_priority: Normal
---


# Private statement

Used at the [module level](../../Glossary/vbe-glossary.md#module-level) to declare private [variables](../../Glossary/vbe-glossary.md#variable) and allocate storage space.

## Syntax

**Private** [ **WithEvents** ] _varname_ [ ( [ _subscripts_ ] ) ] [ **As** [ **New** ] _type_ ]<br/>
[ , [ **WithEvents** ] _varname_ [ ( [ _subscripts_ ] ) ] [ **As** [ **New** ] _type_ ]] **. . .**

<br/>

The **Private** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
|**WithEvents**|Optional. [Keyword](../../Glossary/vbe-glossary.md#keyword) that specifies that _varname_ is an [object variable](../../Glossary/vbe-glossary.md#object-variable) used to respond to events triggered by an [ActiveX object](../../Glossary/vbe-glossary.md#activex-object). **WithEvents** is valid only in [class modules](../../Glossary/vbe-glossary.md#class-module). You can declare as many individual variables as you like by using **WithEvents**, but you can't create [arrays](../../Glossary/vbe-glossary.md#array) with **WithEvents**, nor can you use **New** with **WithEvents**.|
| _varname_|Required. Name of the variable; follows standard variable naming conventions.|
| _subscripts_|Optional. Dimensions of an array variable; up to 60 multiple dimensions may be declared. The _subscripts_ [argument](../../Glossary/vbe-glossary.md#argument) uses the following syntax:<br/><br/>[ _lower_**To** ] _upper_ [ , [ _lower_**To** ] _upper_ ] **. . .**<br/><br/>When not explicitly stated in _lower_, the lower bound of an array is controlled by the **[Option Base](option-base-statement.md)** statement. The lower bound is zero if no **Option Base** statement is present.|
|**New**|Optional. Keyword that enables implicit creation of an object. If you use **New** when declaring the object variable, a new instance of the object is created on first reference to it, so you don't have to use the **[Set](set-statement.md)** statement to assign the object reference. The **New** keyword can't be used to declare variables of any intrinsic [data type](../../Glossary/vbe-glossary.md#data-type). It also can't be used to declare instances of dependent objects, and it can't be used with **WithEvents**.|
| _type_|Optional. Data type of the variable; may be [Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported), [Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (for variable-length strings), **String** _length_ (for fixed-length strings), [Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), or an [object type](../../Glossary/vbe-glossary.md#object-type). Use a separate **As** _type_ clause for each variable being defined.|

## Remarks

**Private** variables are available only to the module in which they are declared.

Use the **Private** statement to declare the data type of a variable. For example, the following statement declares a variable as an **Integer**:

```vb
Private NumberOfEmployees As Integer 

```

You can also use a **Private** statement to declare the object type of a variable. The following statement declares a variable for a new instance of a worksheet:

```vb
Private X As New Worksheet 

```

If the **New** keyword isn't used when declaring an object variable, the variable that refers to the object must be assigned an existing object by using the **Set** statement before it can be used. Until it's assigned an object, the declared object variable has the special value **[Nothing](nothing-keyword.md)**, which indicates that it doesn't refer to any particular instance of an object.

If you don't specify a data type or object type, and there is no **[Deftype](../../concepts/getting-started/deftype-statements.md)** statement in the module, the variable is **Variant** by default.

You can also use the **Private** statement with empty parentheses to declare a dynamic array. After declaring a dynamic array, use the **[ReDim](redim-statement.md)** statement within a procedure to define the number of dimensions and elements in the array. If you try to redeclare a dimension for an array variable whose size was explicitly specified in a **Private**, **[Public](public-statement.md)**, or **[Dim](dim-statement.md)** statement, an error occurs.

When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros. **Variant** variables are initialized to [Empty](../../Glossary/vbe-glossary.md#empty). Each element of a user-defined type variable is initialized as if it were a separate variable.


> [!NOTE] 
> The **Private** statement cannot be used inside a procedure; use the **Dim** statement to declare local variables.


## Example

This example shows the **Private** statement being used at the module level to declare variables as private; that is, they are available only to the module in which they are declared.


```vb
Private Number As Integer ' Private Integer variable. 
Private NameArray(1 To 5) As String ' Private array variable. 
' Multiple declarations, two Variants and one Integer, all Private. 
Private MyVar, YourVar, ThisVar As Integer 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]