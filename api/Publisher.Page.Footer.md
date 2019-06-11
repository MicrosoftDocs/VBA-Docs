---
title: Page.Footer property (Publisher)
keywords: vbapb10.chm393248
f1_keywords:
- vbapb10.chm393248
ms.prod: publisher
api_name:
- Publisher.Page.Footer
ms.assetid: 8ab5a59b-c8d5-6217-098c-c53336ee5311
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.Footer property (Publisher)

Returns a **[HeaderFooter](publisher.headerfooter.md)** object representing the footer of the specified **Page** object. Read-only.


## Syntax

_expression_.**Footer**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

HeaderFooter


## Remarks

This property is only for master pages. A "This feature is only for master pages" error is returned if the **Footer** property is accessed from a **Page** object that is returned form the **Pages** collection. 

A new **HeaderFooter** object is created for the specified master page by accessing this property.


## Example

The following example creates a **HeaderFooter** object and sets it to the footer of the first master page.

```vb
Dim objFooter As HeaderFooter 
Set objFooter = ActiveDocument.MasterPages(1).Footer
```

<br/>

The **HeaderFooter** object returned by the **Footer** property can be used to manipulate the footer content. The following example sets some properties of the **HeaderFooter** object of the first master page.

```vb
With ActiveDocument.masterPages(1) 
 With .Header 
 .TextRange.Text = "Windows" & Chr(13) & "Office" & Chr(13) & "Internet Explorer" 
 With .TextRange.ParagraphFormat 
 .SetListType Value:=pbListTypeBullet, BulletText:="*" 
 .Alignment = pbParagraphAlignmentLeft 
 End With 
 End With 
 With .Footer 
 .TextRange.Hyperlinks.Add Text:=.TextRange, _ 
 Address:="https://www.tailspintoys.com", _ 
 TextToDisplay:="Tailspin" 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]