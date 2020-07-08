---
title: Comments.Add method (PowerPoint)
keywords: vbapp10.chm641004
f1_keywords:
- vbapp10.chm641004
ms.prod: powerpoint
api_name:
- PowerPoint.Comments.Add
ms.assetid: ab520c51-2a8b-2e37-2e4c-8fce7a70a5ab
ms.date: 06/08/2017
localization_priority: Normal
---


# Comments.Add method (PowerPoint)

Returns a **[Comment](PowerPoint.Comment.md)** object that represents a new comment added to a slide.

> [!IMPORTANT]
> **Add()** is now hidden and can't be added in new places. However, **Add()** will continue to work for both modern and legacy comments. On modern comments, we no longer support assigning a modern comment to another author via **Add()**. Instead, modern comments will always be attributed to the signed-in user. If users would like to assign modern comments to another user, they can update their add-in to use **[Add2()](PowerPoint.Comments.Add2.md)**, which was designed for modern comments.

## Syntax

_expression_.**Add** (_Left_, _Top_, _Author_, _AuthorInitials_, _Text_)

_expression_ A variable that represents a **[Comments](PowerPoint.Comments.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the comment, relative to the left edge of the presentation.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the comment, relative to the top edge of the presentation.|
| _Author_|Required|**String**|The author of the comment.|
| _AuthorInitials_|Required|**String**|The author's initials.|
| _Text_|Required|**String**|The comment's text.|

## Return value

Comment


## See also


[Comments Object](PowerPoint.Comments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
