---
title: OLEObjects.Item property (Visio)
keywords: vis_sdr.chm15113765
f1_keywords:
- vis_sdr.chm15113765
ms.prod: visio
api_name:
- Visio.OLEObjects.Item
ms.assetid: a125e2cd-013f-f97a-d4ec-89043cc3bb4b
ms.date: 06/08/2017
localization_priority: Normal
---


# OLEObjects.Item property (Visio)

Returns an item from a collection. The **Item** property is the default property for all collections. Read-only.


## Syntax

_expression_.**Item** (_NameOrIndex_)

_expression_ A variable that represents an **[OLEObjects](Visio.OLEObjects.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NameOrIndex_|Required| **Variant**|Contains the name, unique ID, or index of the object to retrieve.|

## Return value

OLEObject


## Remarks

When retrieving objects from a collection, you can omit **Item** from the expression because it is the default property for all collections. The following statements are equivalent to the syntax example given above:


```vb
objRet = object(index) 
objRet = object(stringExpression) 

```

You can retrieve an object in an **Addons**, **Documents**, **Fonts**, **Hyperlinks**, **Layers**, **Masters**, **MasterShortcuts**, **OLEObjects**, **Pages**, **Shapes**, or **Styles** collection by passing the object's name as a string expression in a **Variant**.

For more information about passing ID strings to the **Item** property, see the topic for the **[UniqueID](Visio.Shape.UniqueID.md)** property in this reference.


> [!NOTE] 
> Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the **Item** property to access an object in the **Masters**, **Pages**, **Shapes**, **Styles**, **Layers**, or **MasterShortcuts** collection by using its local name. Use the **ItemU** property to access an object from one of these collections by using the object's universal name.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]