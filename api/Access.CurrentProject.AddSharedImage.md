---
title: CurrentProject.AddSharedImage method (Access)
keywords: vbaac10.chm14660
f1_keywords:
- vbaac10.chm14660
ms.prod: access
api_name:
- Access.CurrentProject.AddSharedImage
ms.assetid: c6c02f12-6c5f-852a-65b7-a0ffbb3346fd
ms.date: 02/27/2019
localization_priority: Normal
---


# CurrentProject.AddSharedImage method (Access)

Imports the specified image into the database and adds it to the **[SharedResources](Access.SharedResources.md)** collection.


## Syntax

_expression_.**AddSharedImage** (_SharedImageName_, _FileName_)

_expression_ A variable that represents a **[CurrentProject](Access.CurrentProject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SharedImageName_|Required|**String**|Specifies the string used to identify the image in the collection.|
| _FileName_|Required|**String**|Specifies the full name and path to the image file.|

## Remarks

Use the **AddSharedImage** method when you have an image that you want to use repeatedly, such as a company logo. The **AddSharedImage** method makes the image available in the **Insert Image** dropdown of the **Controls** group on the **Design** tab.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]