---
title: Addons.Item property (Visio)
keywords: vis_sdr.chm12513765
f1_keywords:
- vis_sdr.chm12513765
ms.prod: visio
api_name:
- Visio.Addons.Item
ms.assetid: e076376a-d003-7250-2d14-a7ebe1568e75
ms.date: 06/24/2019
localization_priority: Normal
---


# Addons.Item property (Visio)

Returns an item from a collection. The **Item** property is the default property for all collections. Read-only.


## Syntax

_expression_.**Item** (_NameOrIndex_)

_expression_ A variable that represents an **[Addons](Visio.Addons.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NameOrIndex_|Required| **Variant**|Contains the name, unique ID, or index of the object to retrieve.|

## Return value

**[Addon](Visio.Addon.md)**


## Remarks

When retrieving objects from a collection, you can omit **Item** from the expression because it is the default property for all collections. The following statements are equivalent to the syntax shown previously.

```vb
objRet = object(index) 
objRet = object(stringExpression) 

```

You can retrieve an object in an **Addons**, **Documents**, **Fonts**, **Hyperlinks**, **Layers**, **Masters**, **MasterShortcuts**, **OLEObjects**, **Pages**, **Shapes**, or **Styles** collection by passing the object's name as a string expression in a **Variant**.

For more information about passing ID strings to the **Item** property, see the **[Shape.UniqueID](Visio.Shape.UniqueID.md)** property.

> [!NOTE] 
> Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the **Item** property to access an object in the **Masters**, **Pages**, **Shapes**, **Styles**, **Layers**, or **MasterShortcuts** collection by using its local name. Use the **ItemU** property to access an object from one of these collections by using the object's universal name.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]