---
title: Hyperlink.CreateNewDocument Method (PowerPoint)
keywords: vbapp10.chm526012
f1_keywords:
- vbapp10.chm526012
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.CreateNewDocument
ms.assetid: d2de9bbb-a659-3ea3-bdee-244329d88416
ms.date: 06/08/2017
---


# Hyperlink.CreateNewDocument Method (PowerPoint)

Creates a new Web presentation associated with the specified hyperlink.


## Syntax

 _expression_. `CreateNewDocument`( `_FileName_`, `_EditNow_`, `_Overwrite_` )

 _expression_ A variable that represents a [Hyperlink](./PowerPoint.Hyperlink.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The path and file name of the document.|
| _EditNow_|Required|**[MsoTriState](./Office.MsoTriState.md)**|Determines whether the document is opened immediately in its associated editor.|
| _Overwrite_|Required|**[MsoTriState](./Office.MsoTriState.md)**|Determines whether any existing file of the same name in the same folder is overwritten.|

### Return value

Nothing


## Example

This example creates a new Web presentation to be associated with hyperlink one on slide one. The new presentation is called Brittany.ppt, and it overwrites any file of the same name in the HTMLPres folder. The new presentation document is loaded into Microsoft PowerPoint immediately for editing.


```vb
ActivePresentation.Slides(1).Hyperlinks(1).CreateNewDocument _ 
    FileName:="C:\HTMLPres\Brittany.ppt", _ 
    EditNow:=msoTrue, _ 
    Overwrite:=msoTrue
```


## See also


[Hyperlink Object](PowerPoint.Hyperlink.md)

