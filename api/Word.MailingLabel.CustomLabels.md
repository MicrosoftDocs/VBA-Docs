---
title: MailingLabel.CustomLabels property (Word)
keywords: vbawd10.chm152502280
f1_keywords:
- vbawd10.chm152502280
ms.prod: word
api_name:
- Word.MailingLabel.CustomLabels
ms.assetid: c4bad9e7-8da9-d469-4d49-a3b43c5cc4de
ms.date: 06/08/2017
localization_priority: Normal
---


# MailingLabel.CustomLabels property (Word)

Returns a  **[CustomLabels](Word.customlabels.md)** collection that represents the available custom mailing labels. Read-only.


## Syntax

_expression_. `CustomLabels`

_expression_ A variable that represents a '[MailingLabel](Word.MailingLabel.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a new custom label named "AdminAddress" and then creates a page of mailing labels using a predefined return address.


```vb
Dim strAddress As String 
Dim labelNew As CustomLabel 
 
strAddress = "Administration" & vbCr & "Mail Stop 22-16" 
 
Set labelNew = Application.MailingLabel _ 
 .CustomLabels.Add(Name:="AdminAddress", DotMatrix:= False) 
 
With labelNew 
 .Height = InchesToPoints(0.5) 
 .Width = InchesToPoints(1) 
 .HorizontalPitch = InchesToPoints(2.06) 
 .VerticalPitch = InchesToPoints(0.5) 
 .NumberAcross = 4 
 .NumberDown = 20 
 .PageSize = wdCustomLabelLetter 
 .SideMargin = InchesToPoints(0.28) 
 .TopMargin = InchesToPoints(0.5) 
End With 
 
Application.MailingLabel.CreateNewDocument _ 
 Name:="AdminAddress", Address:=strAddress
```


## See also


[MailingLabel Object](Word.MailingLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]