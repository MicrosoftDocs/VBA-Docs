---
title: ParagraphFormat.ListType property (Publisher)
keywords: vbapb10.chm5439521
f1_keywords:
- vbapb10.chm5439521
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListType
ms.assetid: 04ae7157-e864-4e95-74ff-59821eceb286
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.ListType property (Publisher)

Returns a **[PbListType](publisher.pblisttype.md)** constant from the specified **ParagraphFormat** object. Read-only.


## Syntax

_expression_.**ListType**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

PbListType


## Remarks

This property is read-only. To set the **ListType** property of a **ParagraphFormat** object, use the **[SetListType](Publisher.ParagraphFormat.SetListType.md)** method.

The **ListType** property value can be one of the **PbListType** constants declared in the Microsoft Publisher type library.


## Example

This example tests to see if the list type is a numbered list, specifically **pbListTypeArabic**. If the **ListType** property is set to **pbListTypeArabic**, the **ListNumberSeparator** property value is set to **pbListSeparatorParenthesis**.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 
 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]