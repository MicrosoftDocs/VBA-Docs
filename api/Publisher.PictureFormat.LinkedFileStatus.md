---
title: PictureFormat.LinkedFileStatus property (Publisher)
keywords: vbapb10.chm3604787
f1_keywords:
- vbapb10.chm3604787
ms.prod: publisher
api_name:
- Publisher.PictureFormat.LinkedFileStatus
ms.assetid: 43ddffe3-9cc3-b102-c5e8-80f26f63849c
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.LinkedFileStatus property (Publisher)

Returns a **[PbLinkedFileStatus](publisher.pblinkedfilestatus.md)** constant that indicates the status of the file linked to the specified picture. Read-only.


## Syntax

_expression_.**LinkedFileStatus**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

PbLinkedFileStatus


## Remarks

This property only applies to linked picture files. It returns "Permission Denied" for shapes representing embedded or pasted pictures.

Use either of the following properties to determine whether a shape represents a linked picture:

- The **[Type](Publisher.Shape.Type.md)** property of the **Shape** object   
- The **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object
    

The **LinkedFileStatus** property value can be one of the **PbLinkedFileStatus** constants declared in the Microsoft Publisher type library.


## Example

The following example generates a list of the linked pictures in the active publication for which the linked files cannot be found.

```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .LinkedFileStatus = pbLinkedFileMissing Then 
 Debug.Print .Filename 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]