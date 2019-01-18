---
title: Page.Name Property (Visio)
keywords: vis_sdr.chm10913930
f1_keywords:
- vis_sdr.chm10913930
ms.prod: visio
api_name:
- Visio.Page.Name
ms.assetid: 745bb4cf-b79c-4212-325b-40b4e1c9bc81
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Name Property (Visio)

Specifies the name of an object. Read-only.


## Syntax

 _expression_. `Name`

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Return value

String


## Remarks




 **Note**  Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, and layers. When a user names a shape, for example, the user is specifying a local name. Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions, universal names were not visible in the user interface.) 

As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. Use the  **Name** property to get or set a **Hyperlink** , **Layer** , **Master** , **MasterShortcut** , **Page** , **Shape** , **Style** , or **Row** object's local name. Use the **NameU** property to get or set its universal name.


