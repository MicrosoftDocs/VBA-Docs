---
title: Addons object (Visio)
keywords: vis_sdr.chm10035
f1_keywords:
- vis_sdr.chm10035
ms.prod: visio
api_name:
- Visio.Addons
ms.assetid: c58bd4f5-20f6-6eae-d0d2-2ddb6a5a45e6
ms.date: 06/19/2019
localization_priority: Normal
---


# Addons object (Visio)

Represents the set of installed add-ons known to an **[Application](visio.application.md)** object.


## Remarks

To retrieve an **Addons** collection, use the **[Addons](visio.application.addons.md)** property of an **Application** object.

The default property of an **Addons** collection is **Item**.

Installed add-ons are those that Microsoft Visio finds in its **Addons** or **StartUp** paths, those that were installed during the initial setup of Visio, those you have installed by using a Microsoft Windows Installer package, or those that other add-ons have dynamically installed by using the **Add** method of the **Addons** collection.

## Methods

-  [Add](Visio.Addons.Add.md)
-  [GetNames](Visio.Addons.GetNames.md)
-  [GetNamesU](Visio.Addons.GetNamesU.md)

## Properties

-  [Application](Visio.Addons.Application.md)
-  [Count](Visio.Addons.Count.md)
-  [Item](Visio.Addons.Item.md)
-  [ItemU](Visio.Addons.ItemU.md)
-  [ObjectType](Visio.Addons.ObjectType.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]