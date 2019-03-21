---
title: SharedResource object (Access)
keywords: vbaac10.chm14654
f1_keywords:
- vbaac10.chm14654
ms.prod: access
api_name:
- Access.SharedResource
ms.assetid: a97163fa-f833-ed1c-aea5-1a7bab783eba
ms.date: 03/21/2019
localization_priority: Normal
---


# SharedResource object (Access)

Represents a Microsoft Office theme or image that is available as a shared resource in the database.


## Remarks

A shared resource is stored once in the database, but can be used many times. For example, you may want to display your company logo on every form that you create. In earlier versions of Access, you had to import the logo into every form. In Access, you can add the logo as a shared image. It will then be displayed in the **Image Gallery** that appears when you choose the **Insert Image** menu for the **Controls** group on the **Design** tab.

To import an image as a **SharedResource** object, use the **[AddSharedImage](Access.CodeProject.AddSharedImage.md)** method of the **CodeProject** object or the **[AddSharedImage](Access.CurrentProject.AddSharedImage.md)** method of the **CurrentProject** object.


## Methods

- [Delete](Access.SharedResource.Delete.md)

## Properties

- [Name](Access.SharedResource.Name.md)
- [Parent](Access.SharedResource.Parent.md)
- [Type](Access.SharedResource.Type.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]