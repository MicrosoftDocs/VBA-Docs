---
title: PageSetup.AvailablePageSizes property (Publisher)
keywords: vbapb10.chm6946849
f1_keywords:
- vbapb10.chm6946849
ms.prod: publisher
api_name:
- Publisher.PageSetup.AvailablePageSizes
ms.assetid: 5ad79ee6-6d32-6c46-c02e-a9ab252267cf
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.AvailablePageSizes property (Publisher)

Returns the **[PageSizes](publisher.pagesizes.md)** collection that contains all the **[PageSize](Publisher.PageSize.md)** objects available in the current publication.


## Syntax

_expression_.**AvailablePageSizes**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

PageSizes


## Remarks

**PageSize** objects correspond to the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Microsoft Publisher user interface.

Page sizes returned by the **AvailablePageSizes** property include not only the page sizes provided by Microsoft Publisher, but also custom page sizes that you create or download, if any.

As you add or remove custom page sizes, the index number for all existing page sizes may change. 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to create a text file that contains the list of all page sizes available in the current publication and their corresponding index numbers. It saves the text file to the Documents or My Documents folder of the current user.

```vb
Public Sub AvailablePageSizes_Example() 
 
 Dim pubPageSize As Publisher.PageSize 
 Dim pubPageSizes As Publisher.PageSizes 
 Dim intCount As Integer 
 Dim lngPageSizeFile As Long 
 
 intCount = 1 
 
 Set pubPageSizes = ThisDocument.PageSetup.AvailablePageSizes 
 
 lngPageSizeFile = FreeFile 
 Open Environ("USERPROFILE") + "\My Documents\PageSizeList.txt" For Output Access Write As lngPageSizeFile 
 
 For Each pubPageSize In pubPageSizes 
 Write #lngPageSizeFile, pubPageSize.Name, intCount 
 intCount = intCount + 1 
 Next 
 
 Close lngPageSizeFile 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]