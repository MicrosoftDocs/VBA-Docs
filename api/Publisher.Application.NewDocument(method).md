---
title: Application.NewDocument method (Publisher)
keywords: vbapb10.chm131127
f1_keywords:
- vbapb10.chm131127
ms.prod: publisher
api_name:
- Publisher.Application.NewDocument
ms.assetid: 9beb6176-0c46-0ba0-8d41-a9021c624223
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.NewDocument method (Publisher)

Returns a **[Document](publisher.document.md)** object that represents a new publication.


## Syntax

_expression_.**NewDocument** (_Wizard_, _Design_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Wizard_|Optional| **[PbWizard](publisher.pbwizard.md)**|The wizard to use to create the new publication. Can be one of the **PbWizard** constants declared in the Microsoft Publisher type library. The default is **pbWizardNone**.|
|_Design_|Optional| **Long**|The design to apply to the new publication.|

## Return value

Document

## Example

This example creates a new publication and edits the master page to contain a page number in a star in the upper-left corner of the page.

```vb
Sub CreateNewPublication() 
 Dim AppPub As Application 
 Dim DocPub As Document 
 
 Set AppPub = New Publisher.Application 
 Set DocPub = AppPub.NewDocument 
 AppPub.ActiveWindow.Visible = True 
 
 With DocPub.MasterPages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=36, _ 
 Top:=36, Width:=50, Height:=50) 
 .Fill.ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 With .TextFrame.TextRange 
 .InsertPageNumber 
 .ParagraphFormat.Alignment = pbParagraphAlignmentCenter 
 With .Font 
 .Bold = msoTrue 
 .Color.RGB = RGB(Red:=255, Green:=255, Blue:=255) 
 .Size = 12 
 End With 
 End With 
 End With 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]