---
title: Document.WebPagePreview method (Publisher)
keywords: vbapb10.chm196724
f1_keywords:
- vbapb10.chm196724
ms.prod: publisher
api_name:
- Publisher.Document.WebPagePreview
ms.assetid: 44083fae-d21d-9cd3-3553-a4d4346141f5
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.WebPagePreview method (Publisher)

Generates a webpage preview of the specified publication in Internet Explorer.


## Syntax

_expression_.**WebPagePreview**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Remarks

A web preview can be generated for print publications. However, the appearance of the web preview may differ from the printed publication.

The web preview opens with the active page displayed. Preview webpages are generated for each page in the publication. However, if the publication is a print publication or otherwise lacks a navigation bar, there may be no way to navigate to those pages.

Use the **[PublicationType](Publisher.Document.PublicationType.md)** property to determine if a publication is a print publication or a web publication.

This method corresponds to the **Web Page Preview** command on the **File** menu.


## Example

The following example sets the active page of the publication and generates a web preview of the publication.

```vb
 
With ActiveDocument 
 .ActiveView.ActivePage = .Pages(2) 
 .WebPagePreview 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]