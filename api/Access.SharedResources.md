---
title: SharedResources object (Access)
keywords: vbaac10.chm14648
f1_keywords:
- vbaac10.chm14648
ms.prod: access
api_name:
- Access.SharedResources
ms.assetid: 45323141-e7df-1c70-efe2-926c1990d5e0
ms.date: 03/21/2019
localization_priority: Normal
---


# SharedResources object (Access)

Represents the collection of shared resources in the database.


## Remarks

The **SharedResources** collection contains Microsoft Office themes and images that are stored once, but used throughout the database.

For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. It will then be displayed in the **Image Gallery** that appears when you choose the **Insert Image** menu for the **Controls** group on the **Design** tab.

Use the **[Resources](Access.CodeProject.Resources.md)** property of the **CodeProject** object or the **[Resources](Access.CurrentProject.Resources.md)** property of the **CurrentProject** object to enumerate the **SharedResources** collection.

To import an image as a **[SharedResource](Access.SharedResource.md)** object, use the **[AddSharedImage](Access.CodeProject.AddSharedImage.md)** method of the **CodeProject** object or the **[AddSharedImage](Access.CurrentProject.AddSharedImage.md)** method of the **CurrentProject** object.


## Properties

- [Application](Access.SharedResources.Application.md)
- [Count](Access.SharedResources.Count.md)
- [Item](Access.SharedResources.Item.md)
- [Parent](Access.SharedResources.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]