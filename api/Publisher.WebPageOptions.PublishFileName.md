---
title: WebPageOptions.PublishFileName property (Publisher)
keywords: vbapb10.chm544784
f1_keywords:
- vbapb10.chm544784
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.PublishFileName
ms.assetid: d3f52a82-8876-303a-2a73-fdb6dd1ff1cb
ms.date: 06/18/2019
localization_priority: Normal
---


# WebPageOptions.PublishFileName property (Publisher)

Returns or sets a **String** that represents the file name of a webpage (within a web publication) that is being saved as filtered HTML.


## Syntax

_expression_.**PublishFileName**

_expression_ A variable that represents a **[WebPageOptions](Publisher.WebPageOptions.md)** object.


## Return value

String


## Remarks

Specifying a file name for a webpage is optional. When a publication is saved as filtered HTML, Microsoft Publisher automatically generates a file name for any webpage that does not have a file name specified. Use the **[SaveAs](Publisher.Document.SaveAs.md)** method of the **Document** object to save a publication as filtered HTML.

User-defined file names are used only when a publication is saved as filtered HTML.

File names must be specified without a file name extension.

Including invalid characters (such as characters that are not universally allowed in file names that are part of URLs) in the file name generates a run-time error. Invalid characters include: 

- Characters with a code point value of below (decimal) 48.    
- Any double-byte characters.
- The following symbols: `,`, `?`, `>`, `<`, `|`, `:`, and `.`
    
This property corresponds to the **File name** text box in the **Publish to the Web** section of the **Web Page Options** dialog box.


## Example

The following example sets the file name and description of the second page in the active publication. The example assumes that the active publication is a web publication containing at least two pages.

```vb
With ActiveDocument.Pages(2).WebPageOptions 
 .PublishFileName = "CompanyProfile" 
 .Description = "Tailspin Toys Company Profile" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]