---
title: ParagraphFormat.ListNumberSeparator property (Publisher)
keywords: vbapb10.chm5439526
f1_keywords:
- vbapb10.chm5439526
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListNumberSeparator
ms.assetid: 63189011-12a0-c7bc-f6c6-7b17b0dcedf2
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.ListNumberSeparator property (Publisher)

Sets or retrieves a **[PbListSeparator](publisher.pblistseparator.md)** constant that represents the list separator of the specified paragraphs. Read/write.


## Syntax

_expression_.**ListNumberSeparator**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

PbListNumberSeparator


## Remarks

You must set the **[ListType](Publisher.ParagraphFormat.ListType.md)** property to a numbered list type before you set the **ListNumberSeparator** property. Returns an "Access Denied" message if the list is not a numbered list.

The **ListNumberSeparator** property value can be one of the **PbListSeparator** constants.

## Example

This example tests to see if the list type is a numbered list, specifically **pbListTypeArabic** (**[PbListType](publisher.pblisttype.md)** enumeration). If the **ListType** property is set to **pbListTypeArabic**, the **ListNumberSeparator** property is set to **pbListSeparatorParenthesis**. Otherwise, the **[SetListType](Publisher.ParagraphFormat.SetListType.md)** method is called and passes **pbListTypeArabic** as the _Value_ (**PbListType**) parameter, and then the **ListNumberSeparator** property can be set.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 Else 
 .SetListType pbListTypeArabic 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]