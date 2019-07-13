---
title: SlideRange.Comments property (PowerPoint)
keywords: vbapp10.chm532032
f1_keywords:
- vbapp10.chm532032
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Comments
ms.assetid: ff06c024-66cf-d915-e0b0-676b009f93fb
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideRange.Comments property (PowerPoint)

Returns a **[Comments](PowerPoint.Comments.md)** object that represents a collection of comments. Read-only.


## Syntax

_expression_.**Comments**

_expression_ A variable that represents a **[SlideRange](PowerPoint.SlideRange.md)** object.


## Return value

Comments


## Example

The following example adds a comment to a slide.

```vb
Sub AddNewComment()
    ActivePresentation.Slides(1).Comments.Add _
        Left:=0, Top:=0, Author:="John Doe", AuthorInitials:="jd", _
        Text:="Please check this spelling again before the next draft."
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]