---
title: Const directive
keywords: vblr6.chm1014460
f1_keywords:
- vblr6.chm1014460
ms.prod: office
ms.assetid: c5d74b3a-75b1-1263-ab98-82a1a1087207
ms.date: 11/29/2018
localization_priority: Normal
---


# #Const directive

Used to define [conditional compiler constants](../../Glossary/vbe-glossary.md#conditional-compiler-constant) for Visual Basic.

## Syntax

**#Const** _constname_ = _expression_

<br/>

The **#Const** compiler directive syntax has these parts:

|Part|Description|
|:-----|:-----|
| _constname_|Required; **Variant** (**String**). Name of the [constant](../../Glossary/vbe-glossary.md#constant); follows standard [variable](../../Glossary/vbe-glossary.md#variable) naming conventions.|
| _expression_|Required. Literal, other conditional compiler constant, or any combination that includes any or all arithmetic or logical [operators](operator-summary.md) except **Is**.|

## Remarks

Conditional compiler constants are always [Private](../../Glossary/vbe-glossary.md#private) to the [module](../../Glossary/vbe-glossary.md#module) in which they appear. It is not possible to create [Public](../../Glossary/vbe-glossary.md#public) compiler constants by using the **#Const** directive. **Public** compiler constants can only be created in the user interface.

Only conditional compiler constants and literals can be used in _expression_. Using a standard constant defined with **Const**, or using a constant that is undefined, causes an error to occur. Conversely, constants defined by using the **#Const** [keyword](../../Glossary/vbe-glossary.md#keyword) can only be used for conditional compilation.

Conditional compiler constants are always evaluated at the [module level](../../Glossary/vbe-glossary.md#module-level), regardless of their placement in code.

## Example

This example uses the **#Const** directive to declare conditional compiler constants for use in **#If...#Else...#End If** constructs.


```vb
#Const DebugVersion = 1 ' Will evaluate true in #If block. 

```


## See also

- [Const statement](const-statement.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]