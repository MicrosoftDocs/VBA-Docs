---
title: ContainerProperties.AddMember Method (Visio)
keywords: vis_sdr.chm17662355
f1_keywords:
- vis_sdr.chm17662355
ms.prod: visio
api_name:
- Visio.ContainerProperties.AddMember
ms.assetid: fcb97d9f-756e-95fb-8dab-d4aac67862c0
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.AddMember Method (Visio)

Adds a shape or a set of shapes to the container.


## Syntax

 _expression_. `AddMember`( `_pObjectToAdd_` , `_addOptions_` )

 _expression_ A variable that represents a '[ContainerProperties](Visio.ContainerProperties.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pObjectToAdd_|Required| **[UNKNOWN]**|The shape or shapes to add to the container. Can be of type  **[Shape](Visio.Shape.md)** or **[Selection](Visio.Selection.md)**.|
| _addOptions_|Required| **[VisMemberAddOptions](Visio.VisMemberAddOptions.md)**|Determines whether the container should expand to fully contain the added shapes.|

## Return value

 **Nothing**


## Remarks

The  _addOptions_ parameter value must be one of the following **VisMemberAddOptions** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visMemberAddUseResizeSetting**|0|Defer to the setting of the  **[ContainerProperties.ResizeAsNeeded](Visio.ContainerProperties.ResizeAsNeeded.md)** property.|
| **visMemberAddExpandContainer**|1|Expand the container to fit the incoming shapes.|
| **visMemberAddDoNotExpand**|2|Do not expand the container to fit the incoming shapes.|

Passing  **visMemberAddUseResizeSetting** or **visMemberAddDoNotExpand** for _addOptions_ can create a situation in which a shape is a container member but not physically within the container. In such a case, the shape can lose its container membership on subsequent moves or resizing of either the container or the member.

If the container is a list,  **AddMember** adds the specified object to the list container, but not to the list itself. In other words, the shape is contained by the list but is not actually in the list. This is common for shapes in containers that are themselves in a list.

If the  **[ContainerProperties.LockMembership](Visio.ContainerProperties.LockMembership.md)** property is **True** , Microsoft Visio returns a Disabled error.

If the  _pObjectToAdd_ parameter does not contain top-level shapes on the page, Visio returns an Invalid Parameter error.

Visio also returns an Invalid Parameter error if you attempt to use the  **AddMember** method to add the container shape itself or subshapes of the container to the container.

Visio returns an Invalid Target error if  _pObjectToAdd_ does not match the category requirements of the list or the container. Shapes can be assigned categories, and containers can have required and excluded categories.

Categories are user-defined strings that you can use to categorize shapes and, thereby, to restrict membership in a container. You can define categories in the User.msvShapeCategories cell in the ShapeSheet for a shape. You can define multiple categories for a shape by separating them with semicolons.


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **AddMember** method to add a new member (vsoShape) to an existing container (vsoContainerShape) on a page. The code assumes that vsoShape already overlaps vsoContainerShape.


```vb
vsoContainerShape.ContainerProperties.AddMember vsoShape, visMemberAddExpandContainer
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]