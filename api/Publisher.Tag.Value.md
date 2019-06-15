---
title: Tag.Value property (Publisher)
keywords: vbapb10.chm4718596
f1_keywords:
- vbapb10.chm4718596
ms.prod: publisher
api_name:
- Publisher.Tag.Value
ms.assetid: dee3b69b-ae5b-df13-561e-84105057979a
ms.date: 06/15/2019
localization_priority: Normal
---


# Tag.Value property (Publisher)

Returns or sets a **Variant** that represents the value of a tag of a shape, page, or publication. Read/write.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[Tag](Publisher.Tag.md)** object.


## Example

This example creates a new tag for the active publication and then displays the value of the tag.

```vb
Sub CreatePublicationTag() 
 With ActiveDocument 
 .Tags.Add Name:="ActivePub", Value:="This is the active publication." 
 MsgBox .Tags(1).Value 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]