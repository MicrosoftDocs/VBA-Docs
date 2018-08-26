---
title: Const Directive
keywords: vblr6.chm1014460
f1_keywords:
- vblr6.chm1014460
ms.prod: office
ms.assetid: c5d74b3a-75b1-1263-ab98-82a1a1087207
ms.date: 06/08/2017
---


# #Const Directive

<<<<<<< HEAD
Used to define [conditional compiler constants](../../Glossary/vbe-glossary.md) for Visual Basic.
=======
Used to define [conditional compiler constants](../../Glossary/vbe-glossary.md#conditional-compiler-constant) for Visual Basic.
>>>>>>> master

## Syntax

 **#Const** _constname_ = _expression_

The **#Const** compiler directive syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
| _constname_|Required;  **Variant** (**String**). Name of the [constant](../../Glossary/vbe-glossary.md); follows standard [variable](../../Glossary/vbe-glossary.md) naming conventions.|
| _expression_|Required. Literal, other conditional compiler constant, or any combination that includes any or all arithmetic or logical operators except  **Is**.|

 **Remarks**
Conditional compiler constants are always [Private](../../Glossary/vbe-glossary.md) to the [module](../../Glossary/vbe-glossary.md) in which they appear. It is not possible to create [Public](../../Glossary/vbe-glossary.md) compiler constants using the **#Const** directive. **Public** compiler constants can only be created in the user interface.
Only conditional compiler constants and literals can be used in  _expression_. Using a standard constant defined with **Const**, or using a constant that is undefined, causes an error to occur. Conversely, constants defined using the **#Const** [keyword](../../Glossary/vbe-glossary.md) can only be used for conditional compilation.
Conditional compiler constants are always evaluated at the [module level](../../Glossary/vbe-glossary.md), regardless of their placement in code.
=======
| _constname_|Required;  **Variant** (**String**). Name of the [constant](../../Glossary/vbe-glossary.md#constant); follows standard [variable](../../Glossary/vbe-glossary.md#variable) naming conventions.|
| _expression_|Required. Literal, other conditional compiler constant, or any combination that includes any or all arithmetic or logical operators except  **Is**.|

## Remarks

Conditional compiler constants are always [Private](../../Glossary/vbe-glossary.md#private) to the [module](../../Glossary/vbe-glossary.md#module) in which they appear. It is not possible to create [Public](../../Glossary/vbe-glossary.md#public) compiler constants using the **#Const** directive. **Public** compiler constants can only be created in the user interface.
Only conditional compiler constants and literals can be used in  _expression_. Using a standard constant defined with **Const**, or using a constant that is undefined, causes an error to occur. Conversely, constants defined using the **#Const** [keyword](../../Glossary/vbe-glossary.md#keyword) can only be used for conditional compilation.
Conditional compiler constants are always evaluated at the [module level](../../Glossary/vbe-glossary.md#module-level), regardless of their placement in code.
>>>>>>> master

## Example

This example uses the  **#Const** directive to declare conditional compiler constants for use in **#If...#Else...#End If** constructs.


```vb
#Const DebugVersion = 1 ' Will evaluate true in #If block. 

```


