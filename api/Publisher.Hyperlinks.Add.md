---
title: Hyperlinks.Add method (Publisher)
keywords: vbapb10.chm6881284
f1_keywords:
- vbapb10.chm6881284
ms.prod: publisher
api_name:
- Publisher.Hyperlinks.Add
ms.assetid: f5a8cc01-a571-623d-bfab-fe48e43a21b1
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlinks.Add method (Publisher)

Adds a new **[Hyperlink](Publisher.Hyperlink.md)** object to the specified **Hyperlinks** collection and returns the new **Hyperlink** object.


## Syntax

_expression_.**Add** (_Text_, _Address_, _RelativePage_, _PageID_, _TextToDisplay_)

_expression_ A variable that represents a **[Hyperlinks](Publisher.Hyperlinks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Text_|Required| **TextRange**| **TextRange** object. The text range to be converted into a hyperlink.|
|_Address_|Optional| **String**|The address of the new hyperlink. If _RelativePage_ is **pbHlinkTargetTypeURL** (default) or **pbHlinkTargetTypeEmail**, _Address_ must be specified, or an error occurs.|
|_RelativePage_|Optional| **[PbHlinkTargetType](publisher.pbhlinktargettype.md)**| The type of hyperlink to add. Can be one of the **PbHlinkTargetType** constants; the default is **pbHlinkTargetTypeURL**.|
|_PageID_|Optional| **Long**|The page ID of the destination page for the new hyperlink. If _RelativePage_ is **pbHlinkTargetTypePageID**, _PageID_ must be specified, or an error occurs. The page ID corresponds to the **[PageID](Publisher.Hyperlink.PageID.md)** property of the destination page.|
|_TextToDisplay_|Optional| **String**|The display text of the new hyperlink. If specified, _TextToDisplay_ replaces the text range specified by the _Text_ argument.|

## Return value

Hyperlink


## Example

The following example adds hyperlinks to shape one and shape two on page one of the active publication. The first hyperlink points to an external website, and the second link points to the fourth page in the publication. Shape one and shape two must be text boxes, and there must be at least four pages in the publication for this example to work.

```vb
Dim hypNew As Hyperlink 
Dim lngPageID As Long 
Dim strPage As String 
 
With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 Address:="https://www.tailspintoys.com/", _ 
 TextToDisplay:="Tailspin") 
End With 
 
lngPageID = ActiveDocument.Pages(4).PageID 
strPage = "Go to page " _ 
 & Str(ActiveDocument.Pages(4).PageNumber) 
 
With ActiveDocument.Pages(1).Shapes(2).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 RelativePage:=pbHlinkTargetTypePageID, _ 
 PageID:=lngPageID, _ 
 TextToDisplay:=strPage) 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]