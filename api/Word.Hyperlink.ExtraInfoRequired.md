---
title: Hyperlink.ExtraInfoRequired property (Word)
keywords: vbawd10.chm161285105
f1_keywords:
- vbawd10.chm161285105
ms.prod: word
api_name:
- Word.Hyperlink.ExtraInfoRequired
ms.assetid: 066a4dbf-f5ea-f708-cd57-f8e515a258d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.ExtraInfoRequired property (Word)

 **True** if extra information is required to resolve the specified hyperlink. Read-only **Boolean**.


## Syntax

_expression_. `ExtraInfoRequired`

_expression_ A variable that represents a '[Hyperlink](Word.Hyperlink.md)' object.


## Remarks

You can specify extra information by using the ExtraInfo argument with the  **[Follow](Word.Hyperlink.Follow.md)** or **[FollowHyperlink](Word.Document.FollowHyperlink.md)** method. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name.


## Example

This example inserts a hyperlink to www.msn.com and then follows the hyperlink if extra information isn't required.


```vb
Dim hypTemp As Hyperlink 
 
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter "MSN " 
 .Previous 
End With 
Set hypTemp = ActiveDocument.Hyperlinks.Add( _ 
 Address:="https://www.msn.com", _ 
 Anchor:=Selection.Range) 
If hypTemp.ExtraInfoRequired = False Then 
 hypTemp.Follow 
End If
```


## See also


[Hyperlink Object](Word.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]