---
title: ParagraphFormat.SetListType method (Publisher)
keywords: vbapb10.chm5439520
f1_keywords:
- vbapb10.chm5439520
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.SetListType
ms.assetid: 6900aac5-fb3f-5813-309c-1422d38c8301
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.SetListType method (Publisher)

Sets the list type of the specified **ParagraphFormat** object. 


## Syntax

_expression_.**SetListType** (_Value_, _BulletText_)

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Value_|Required| **[PbListType](publisher.pblisttype.md)**|Represents the list type of the specified **ParagraphFormat** object. Can be one of the **PbListType** constants declared in the Microsoft Publisher type library. |
|_BulletText_|Optional| **String**| A string that represents the text of the list bullet.|


## Remarks

If _Value_ is a bulleted list and the _BulletText_ parameter is missing, the first bullet from the **Bullets and Numbering** dialog box is used.

_BulletText_ is limited to one character. 

A run-time error occurs if the _BulletText_ parameter is provided and the _Value_ parameter is not set to **pbListTypeBullet**.


## Example

This example tests to see if the list type is a numbered list, specifically **pbListTypeArabic**. If the **[ListType](Publisher.ParagraphFormat.ListType.md)** property is set to **pbListTypeArabic**, the **[ListNumberSeparator](Publisher.ParagraphFormat.ListNumberSeparator.md)** property is set to **pbListSeparatorParenthesis**. Otherwise, the **SetListType** method is called and passes **pbListTypeArabic** as the _Value_ parameter, and then the **ListNumberSeparator** property can be set.

```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeArabic Then 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 Else 
 .SetListType pbListTypeArabic 
 .ListNumberSeparator = pbListSeparatorParenthesis 
 End If 
End With 
 
```

<br/>

This example demonstrates how an organized document structure containing named text frames with lists can be configured. This example assumes that the publication has a naming convention for **[TextFrame](publisher.textframe.md)** objects containing lists that use the word "list" as a prefix. This example uses nested collection iterations to access each of the **TextFrame** objects in each **Shapes** collection of each **Page**. The **ParagraphFormat** object of each **TextFrame** name with the prefix "list" has the **ListType** and **ListBulletFontSize** properties set.

```vb
Dim objPage As page 
Dim objShp As Shape 
Dim objTxtFrm As TextFrame 
 
'Iterate through all pages of th ePublication 
For Each objPage In ActiveDocument.Pages 
 'Iterate through the Shapes collection of objPage 
 For Each objShp In objPage.Shapes 
 'Find each TextFrame object 
 If objShp.Type = pbTextFrame Then 
 'If the name of the TextFrame begins with "list" 
 If InStr(1, objShp.Name, "list") <> 0 Then 
 Set objTxtFrm = objShp.TextFrame 
 With objTxtFrm 
 With .TextRange 
 With .ParagraphFormat 
 .SetListType pbListTypeBullet, "*" 
 .ListBulletFontSize = 24 
 End With 
 End With 
 End With 
 End If 
 End If 
 Next 
Next 
 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]