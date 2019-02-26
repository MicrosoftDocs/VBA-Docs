---
title: CodeProject.AddSharedImage method (Access)
keywords: vbaac10.chm14660
f1_keywords:
- vbaac10.chm14660
ms.prod: access
api_name:
- Access.CodeProject.AddSharedImage
ms.assetid: 7e1e0455-65e0-820e-e25c-17989a40000b
ms.date: 02/27/2019
localization_priority: Normal
---


# CodeProject.AddSharedImage method (Access)

Imports the specified image into the database and adds it to the **[SharedResources](Access.SharedResources.md)** collection.


## Syntax

_expression_.**AddSharedImage** (_SharedImageName_, _FileName_)

_expression_ A variable that represents a **[CodeProject](Access.CodeProject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SharedImageName_|Required|**String**|Specifies the string used to identify the image in the collection.|
| _FileName_|Required|**String**|Specifies the full name and path to the image file.|

## Remarks

Use the **AddSharedImage** method when you have an image that you want to use repeatedly, such as a company logo. The **AddSharedImage** method makes the image available in the **Insert Image** dropdown of the **Controls** group on the **Design** tab.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]