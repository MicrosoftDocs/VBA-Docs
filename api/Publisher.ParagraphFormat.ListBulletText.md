---
title: ParagraphFormat.ListBulletText property (Publisher)
keywords: vbapb10.chm5439523
f1_keywords:
- vbapb10.chm5439523
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListBulletText
ms.assetid: fa80957a-be91-398f-a24f-5a0449a9466f
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.ListBulletText property (Publisher)

Returns a **String** representing the list bullet text from the specified paragraphs. Read-only.


## Syntax

_expression_.**ListBulletText**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

String


## Remarks

The **ListBulletText** property is limited to one character.

This property is read-only. To set the **ListBulletText** property of a bulleted list, use the **[SetListType](Publisher.ParagraphFormat.SetListType.md)** method.

Returns an "Access Denied" message if the list is not a bulleted list.


## Example

This example tests to see if the list type is a bulleted list. If it is, a test is made to see that the list bullet text is set to `"*"`. If it is not, the **SetListType** method is called and passes **pbListTypeBullet** as the _Value_ (**PbListType**) parameter and `"*"` as the _BulletText_ parameter.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 If Not .ListBulletText = "*" Then 
 .SetListType pbListTypeBullet, "*" 
 End If 
 End If 
End With 
 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]