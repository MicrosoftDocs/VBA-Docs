---
title: PictureFormat.Replace method (Publisher)
keywords: vbapb10.chm3604786
f1_keywords:
- vbapb10.chm3604786
ms.prod: publisher
api_name:
- Publisher.PictureFormat.Replace
ms.assetid: b2bce79a-5c46-1473-601d-a4a25176edeb
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.Replace method (Publisher)

Replaces the specified picture. Returns **Nothing**.


## Syntax

_expression_.**Replace** (_PathName_, _InsertAs_)

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PathName_|Required| **String**|The name and path of the file with which you want to replace the specified picture.|
|_InsertAs_|Optional| **[PbPictureInsertAs](publisher.pbpictureinsertas.md)**|The manner in which you want the picture file inserted into the document: linked or embedded. Can be one of the **PbPictureInsertAs** constants declared in the Microsoft Publisher type library; the default value is **pbPictureInsertAsOriginalState**.|

## Remarks

Use the **Replace** method to update linked picture files that have been modified since they were inserted into the document. 

Use the **[LinkedFileStatus](Publisher.PictureFormat.LinkedFileStatus.md)** property to determine if a linked picture has been modified.


## Example

The following example replaces every occurrence of a specific picture in the active publication with another picture.

```vb
Sub ReplaceLogo() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strExistingArtName As String 
Dim strReplaceArtName As String 
 
 
strExistingArtName = "C:\path\logo 1.bmp" 
strReplaceArtName = "C:\path\logo 2.bmp" 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .Filename = strExistingArtName Then 
 .Replace (strReplaceArtName) 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

<br/>

This example tests each linked picture to determine if the linked file has been modified since it was inserted into the publication. If it has, the picture is updated by replacing the file with itself.

```vb
Sub UpdateModifiedLinkedPictures() 
 
Dim pgLoop As Page 
Dim shpLoop As Shape 
Dim strPictureName As String 
 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .LinkedFileStatus = pbLinkedFileModified Then 
 strPictureName = .Filename 
 .Replace (strPictureName) 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]