---
title: UIObject object (Visio)
keywords: vis_sdr.chm10300
f1_keywords:
- vis_sdr.chm10300
ms.prod: visio
api_name:
- Visio.UIObject
ms.assetid: 2d842398-df53-0d59-6ee5-89d411440863
ms.date: 06/19/2019
localization_priority: Normal
---


# UIObject object (Visio)

Represents a set of Microsoft Visio menus, toolbars, and accelerators, from either the built-in Visio user interface or a customized version of it. 

> [!NOTE] 
> Starting with Visio 2010, the Microsoft Office Fluent user interface (UI) replaced the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

## Remarks

To retrieve a **UIObject** object that contains Visio menus and accelerators, use the **[BuiltInMenus](visio.application.builtinmenus.md)** property of an **Application** object and then the **MenuSets** or **AccelTables** collections of the **UIObject** object returned from the **BuiltInMenus** property.
    
To retrieve a **UIObject** object that contains Visio toolbars, use the **[BuiltInToolbars](visio.application.builtintoolbars.md)** property of an **Application** object and then the **ToolbarSets** collection of the **UIObject** object returned from the **BuiltInToolbars** property.
    
If an **Application** object or **[Document](visio.document.md)** object has a customized user interface, use the **CustomMenus** or **CustomToolbars** properties to retrieve **UIObject** objects that represent these.

A **UIObject** object can be stored in a file and loaded into Visio. Use the **SaveToFile** method to save the object and the **LoadFromFile** method to load it, or set the **CustomMenusFile** or **CustomToolbarsFile** property of an **Application** object or **Document** object to the name of the stored user interface file.

Beginning with Visio 2002, a program can manipulate menus and toolbars in the Visio user interface by manipulating the **CommandBars** collection returned by the **CommandBars** property. The **CommandBars** collection has an interface identical to the **CommandBars** collection exposed by the suite of Microsoft System applications such as Microsoft Word and Microsoft Excel. Consequently, programs can manipulate the Visio menus and toolbars by using either the **CommandBars** collection or **UIObject** objects.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).

## Methods

-  [LoadFromFile](Visio.UIObject.LoadFromFile.md)
-  [SaveToFile](Visio.UIObject.SaveToFile.md)
-  [UpdateUI](Visio.UIObject.UpdateUI.md)

## Properties

-  [AccelTables](Visio.UIObject.AccelTables.md)
-  [Clone](Visio.UIObject.Clone.md)
-  [DisplayKeysInTooltips](Visio.UIObject.DisplayKeysInTooltips.md)
-  [DisplayTooltips](Visio.UIObject.DisplayTooltips.md)
-  [LargeButtons](Visio.UIObject.LargeButtons.md)
-  [MenuAnimationStyle](Visio.UIObject.MenuAnimationStyle.md)
-  [MenuSets](Visio.UIObject.MenuSets.md)
-  [Name](Visio.UIObject.Name.md)
-  [ToolbarSets](Visio.UIObject.ToolbarSets.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]