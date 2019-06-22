---
title: Windows.ItemFromID property (Visio)
keywords: vis_sdr.chm11713775
f1_keywords:
- vis_sdr.chm11713775
ms.prod: visio
api_name:
- Visio.Windows.ItemFromID
ms.assetid: 19049ae8-b070-3da7-ce6a-446e547b4d5d
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows.ItemFromID property (Visio)

Returns an item of a collection using the ID of the item. Read-only.


## Syntax

_expression_.**ItemFromID** (_nID_)

_expression_ A variable that represents a **[Windows](Visio.Windows.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _nID_|Required| **Long**|The ID of the object to retrieve.|

## Return value

Window


## Remarks

The ID of a **Shape** object uniquely identifies the shape within its page or master.

The ID of a **Style** object uniquely identifies the style within its document.

The ID of a **Font** object corresponds to the number stored in the Font cell of a row in a shape's Character Properties section. The ID associated with a particular font varies between systems or as fonts are installed on and removed from a given system.

The ID of an **Event** object uniquely identifies an event in its **EventList** collection for the life of the collection.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this property maps to the following types:


- **Microsoft.Office.Interop.Visio.IVWindows.get_ItemFromID(int)**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]