---
title: ContactItem.SaveBusinessCardImage method (Outlook)
keywords: vbaol11.chm1096
f1_keywords:
- vbaol11.chm1096
ms.prod: outlook
api_name:
- Outlook.ContactItem.SaveBusinessCardImage
ms.assetid: 889728f2-2c17-6b83-a858-bb32ef5845e6
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.SaveBusinessCardImage method (Outlook)

Saves an image of the business card generated from the specified  **[ContactItem](Outlook.ContactItem.md)** object.


## Syntax

_expression_. `SaveBusinessCardImage`( `_Path_` )

 _expression_ An expression that returns a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The fully qualified path and file name of the image to be saved.|

## Remarks

This method generates an image, as a Portable Network Graphics (.png) file, of the business card generated from the specified  **ContactItem** object. If the path and file name specified in Path cannot be resolved, an error occurs.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]