---
title: Shapes.ItemFromID property (Visio)
keywords: vis_sdr.chm11313775
f1_keywords:
- vis_sdr.chm11313775
ms.prod: visio
api_name:
- Visio.Shapes.ItemFromID
ms.assetid: 0e8e80a2-94f0-f451-b914-f8d8a56a3ef2
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.ItemFromID property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

_expression_.**ItemFromID** (_nID_)

_expression_ A variable that represents a **[Shapes](Visio.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

## Return value

Shape


## Remarks

The ID of a **Shape** object uniquely identifies the shape within its page or master. You can determine the ID of a shape by displaying the **Special** dialog box (select the shape, and then click **Shape Name** on the [Developer](../visio/How-to/run-visio-in-developer-mode.md) tab.)

The ID of a **Style** object uniquely identifies the style within its document.

The ID of a **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this property maps to the following types:


- **Microsoft.Office.Interop.Visio.IVShapes.get_ItemFromID**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]