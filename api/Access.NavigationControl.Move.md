---
title: NavigationControl.Move method (Access)
keywords: vbaac10.chm11144
f1_keywords:
- vbaac10.chm11144
ms.prod: access
api_name:
- Access.NavigationControl.Move
ms.assetid: bbf4e87e-8468-7cfd-7cd4-5f423a6517c8
ms.date: 06/08/2017
localization_priority: Normal
---


# NavigationControl.Move method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

_expression_. `Move`( ` _Left_`, ` _Top_`, ` _Width_`, ` _Height_` )

_expression_ A variable that represents a [NavigationControl](Access.NavigationControl.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Variant**||
| _Top_|Optional|**Variant**||
| _Width_|Optional|**Variant**||
| _Height_|Optional|**Variant**||

## Remarks

Only the  _Left_ argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specify _Width_ without specifying _Left_ and _Top_. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


[NavigationControl Object](Access.NavigationControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]