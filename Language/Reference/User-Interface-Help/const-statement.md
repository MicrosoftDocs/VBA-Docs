---
title: Const Statement
keywords: vblr6.chm1008877
f1_keywords:
- vblr6.chm1008877
ms.prod: office
ms.assetid: 99e2d1e1-ed30-77d3-3366-6438e9373308
ms.date: 06/08/2017
---


# Const Statement

Declares [constants](../../Glossary/vbe-glossary.md#constant) for use in place of literal values.

## Syntax

[ **Public** | **Private** ] **Const** _constname_ [ **As** _type_ ] **=** _expression_

The  **Const** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. [Keyword](../../Glossary/vbe-glossary.md#Keyword) used at [module level](../../Glossary/vbe-glossary.md#module-level) to declare constants that are available to all [procedures](../../Glossary/vbe-glossary.md#procedure) in all [modules](../../Glossary/vbe-glossary.md#module). Not allowed in procedures.|
|**Private**|Optional. Keyword used at module level to declare constants that are available only within the module where the [declaration](../../Glossary/vbe-glossary.md#declaration) is made. Not allowed in procedures.|
| _constname_|Required. Name of the constant; follows standard [variable](../../Glossary/vbe-glossary.md#variable) naming conventions.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#Data-type) of the constant; may be [Byte](../../Glossary/vbe-glossary.md#Byte), [Boolean](../../Glossary/vbe-glossary.md#Boolean), [Integer](../../Glossary/vbe-glossary.md#Integer), [Long](../../Glossary/vbe-glossary.md#Long), [Currency](../../Glossary/vbe-glossary.md#Currency), [Single](../../Glossary/vbe-glossary.md#Single), [Double](../../Glossary/vbe-glossary.md#Double), [Decimal](../../Glossary/vbe-glossary.md#Decimal) (not currently supported), [Date](../../Glossary/vbe-glossary.md#Date), [String](../../Glossary/vbe-glossary.md#String), or [Variant](../../Glossary/vbe-glossary.md#Variant). Use a separate  **As** _type_ clause for each constant being declared.|
| _expression_|Required. Literal, other constant, or any combination that includes all arithmetic or logical operators except  **Is**.|

## Remarks

Constants are private by default. Within procedures, constants are always private; their visibility can't be changed. In [standard modules](../../Glossary/vbe-glossary.md#standard-module), the default visibility of module-level constants can be changed using the  **Public** keyword. In [class modules](../../Glossary/vbe-glossary.md#class-module), however, constants can only be private and their visibility can't be changed using the  **Public** keyword.
To combine several constant declarations on the same line, separate each constant assignment with a comma. When constant declarations are combined in this way, the  **Public** or **Private** keyword, if used, applies to all of them.
You can't use variables, user-defined functions, or intrinsic Visual Basic functions (such as  **Chr**) in [expressions](../../Glossary/vbe-glossary.md#expression) assigned to constants.

 **Note**  Constants can make your programs self-documenting and easy to modify. Unlike variables, constants can't be inadvertently changed while your program is running.

If you don't explicitly declare the constant type using  **As** _type_, the constant has the data type that is most appropriate for _expression_.
Constants declared in a  **Sub**, **Function**, or **Property** procedure are local to that procedure. A constant declared outside a procedure is defined throughout the module in which it is declared. You can use constants anywhere you can use an expression.

## Example

This example uses the  **Const** statement to declare constants for use in place of literal values. **Public** constants are declared in the General section of a standard module, rather than a class module. **Private** constants are declared in the General section of any type of module.


```vb
' Constants are Private by default. 
Const MyVar = 459 
 
' Declare Public constant. 
Public Const MyString = "HELP" 
 
' Declare Private Integer constant. 
Private Const MyInt As Integer = 5 
 
' Declare multiple constants on same line. 
Const MyStr = "Hello", MyDouble As Double = 3.4567 

```


