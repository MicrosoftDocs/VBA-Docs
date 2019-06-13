---
title: Shapes.AddGroupWizard method (Publisher)
keywords: vbapb10.chm2162727
f1_keywords:
- vbapb10.chm2162727
ms.prod: publisher
api_name:
- Publisher.Shapes.AddGroupWizard
ms.assetid: 5a84f055-7f30-0757-f507-40ee34b214f4
ms.date: 06/14/2019
localization_priority: Normal
---


# Shapes.AddGroupWizard method (Publisher)

Adds a **[Shape](Publisher.Shape.md)** object representing a Design Gallery object to the publication.


## Syntax

_expression_.**AddGroupWizard** (_Wizard_, _Left_, _Top_, _Width_, _Height_, _Design_)

_expression_ A variable that represents a **[Shapes](Publisher.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Wizard_|Required| **[PbWizardGroup](Publisher.PbWizardGroup.md)**|The type of Design Gallery object to add to the publication. Can be one of the **PbWizardGroup** constants declared in the Microsoft Publisher type library.|
|_Left_ |Required| **Variant**|The position of the Design Gallery object's left edge relative to the left edge of the page, measured in points.|
|_Top_ |Required| **Variant**|The position of the Design Gallery object's top edge relative to the top edge of the page, measured in points.|
|_Width_|Optional| **Variant**|The width of the new Design Gallery object.|
|_Height_|Optional| **Variant**|The height of the new Design Gallery object.|
|_Design_|Optional| **Long**|The design of the object to be added.|

## Return value

Shape


## Example

This example adds a web table of contents to the active publication.

```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddGroupWizard Wizard:=pbWizardGroupTableOfContents, _ 
 Left:=100, Top:=100
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]