---
title: Shapes.FindShapeByWizardTag method (Publisher)
keywords: vbapb10.chm2162728
f1_keywords:
- vbapb10.chm2162728
ms.prod: publisher
api_name:
- Publisher.Shapes.FindShapeByWizardTag
ms.assetid: f1018f3a-4f8f-2686-ac58-6eee8827c743
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.FindShapeByWizardTag method (Publisher)

Returns a **[ShapeRange](Publisher.ShapeRange.md)** object representing one or all of the shapes placed in a publication by a wizard and bearing the specified wizard tag.


## Syntax

_expression_.**FindShapeByWizardTag** (_WizardTag_, _Instance_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_WizardTag_|Required| **[PbWizardTag](Publisher.PbWizardTag.md)**|Specifies the wizard tag for which to search. Can be one of the **PbWizardTag** constants declared in the Microsoft Publisher type library.|
|_Instance_|Optional| **Long**|Specifies which instance of a shape with the specified wizard tag is returned. For _Instance_ equal to n, the nth instance of a shape with the specified wizard tag is returned. If no value for _Instance_ is specified, all the shapes with the specified wizard tag are returned.|

## Return value

ShapeRange


## Example

The following example finds the second instance of a shape with the wizard tag **pbWizardDate** and assigns it to a variable.

```vb
Dim shpWizardTag As Shape 
 
Set shpWizardTag = ActiveDocument.FindShapeByWizardTag(WizardTag:=pbWizardDate, Instance:=2)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]