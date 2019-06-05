---
title: Document.SaveAs method (Publisher)
keywords: vbapb10.chm196696
f1_keywords:
- vbapb10.chm196696
ms.prod: publisher
api_name:
- Publisher.Document.SaveAs
ms.assetid: ba8b85d7-8ca9-dcf5-12b4-4cabced743e6
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.SaveAs method (Publisher)

Saves the specified publication with a new name or format.


## Syntax

_expression_.**SaveAs** (_FileName_, _Format_, _AddToRecentFiles_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_ |Optional| **Variant**|The name for the publication. The default is the current folder and file name. If the publication has never been saved, the default name is used, for example, Publication1.pub. If a publication with the specified file name already exists, the publication is overwritten without the user being prompted first.|
|_Format_ |Optional| **[PbFileFormat](publisher.pbfileformat.md)**|The format in which the publication is saved. Can be one of the **PbFileFormat** constants declared in the Microsoft Publisher type library. The default is **pbFilePublication**.|
|_AddToRecentFiles_ |Optional| **Boolean**| **True** to add the publication to the list of recently used files on the **File** menu. Default is **True**.|

## Remarks

If there is insufficient memory or disk space to save the file, an error occurs.

Calling the **SaveAs** method always performs saves in the foreground regardless of whether background saves are enabled.


## Example

This example saves the active publication as a Microsoft Publisher 2000 file.

```vb
ActiveDocument.SaveAs FileName:="ReportPub2000.pub", Format:=pbFilePublisher2000
```

<br/>

This example saves the active publication as Test.rtf in Rich Text Format (RTF).

```vb
ActiveDocument.SaveAs FileName:="Test.rtf", Format:=pbFileRTF
```

<br/>

This example saves the active web publication as a set of filtered HTML pages and supporting files. Note that the .htm extension is automatically added to the value of the _FileName_ parameter if the value of the _Format_ parameter is **pbFileHTMLFiltered**.

```vb
With ActiveDocument 
 .SaveAs Filename:="CompanyContacts", Format:=pbFileHTMLFiltered 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]