---
title: Calling property procedures (VBA)
keywords: vbcn6.chm1101365
f1_keywords:
- vbcn6.chm1101365
ms.prod: office
ms.assetid: 37dfc0de-5db0-85bd-0c15-6d876b6abff9
ms.date: 12/21/2018
localization_priority: Normal
---


# Calling property procedures

The following table lists the syntax for calling property procedures:

|Property procedure|Syntax|
|:-----|:-----|
|**[Property Let](../../reference/user-interface-help/property-let-statement.md)**|[ _object_.] _propname_ (_arguments_)] = _argument_|
|**[Property Get](../../reference/user-interface-help/property-get-statement.md)**| _varname_ = [ _object_.] _propname_ (_arguments_)]|
|**[Property Set](../../reference/user-interface-help/property-set-statement.md)**|**Set** [ _object_.] _propname_. [ (_arguments_) ] = _varname_|

When you call a **Property Let** or **Property Set** procedure, one [argument](../../Glossary/vbe-glossary.md#argument) always appears on the right side of the equal sign (**=**).

When you declare a **Property Let** or **Property Set** procedure with multiple arguments, Visual Basic passes the argument on the right side of the call to the last argument in the **Property Let** or **Property Set** declaration. 

For example, the following diagram shows how arguments in the property procedure call relate to arguments in the **Property Let** declaration:

![Property Let](../../../images/abhlp002_ZA01201812.gif)

In practice, the only use for property procedures with multiple arguments is to create [arrays](../../Glossary/vbe-glossary.md#array) of [properties](../../Glossary/vbe-glossary.md#property).

## See also

- [Visual Basic conceptual topics](../../reference/user-interface-help/visual-basic-conceptual-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
