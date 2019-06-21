---
title: Shape.HasCategory method (Visio)
keywords: vis_sdr.chm11262250
f1_keywords:
- vis_sdr.chm11262250
ms.prod: visio
api_name:
- Visio.Shape.HasCategory
ms.assetid: 91115794-31ab-73b1-d1ec-ca249a57a61f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.HasCategory method (Visio)

Returns  **True** if the specified category is in the shape categories list.


## Syntax

_expression_. `HasCategory`( `_Category_` )

_expression_ A variable that represents a **[Shape](Visio.Shape.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Category_|Required| **String**|The name of the category.|

## Return value

 **Boolean**


## Remarks

You can pass only a single category for the  _Category_ parameter. Passing a semicolon-delimited list of categories produces an Invalid Parameter error.

Categories are user-defined strings that you can use to categorize shapes and thereby to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating those categories with semi-colons.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]