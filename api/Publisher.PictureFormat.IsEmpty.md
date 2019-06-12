---
title: PictureFormat.IsEmpty property (Publisher)
keywords: vbapb10.chm3604788
f1_keywords:
- vbapb10.chm3604788
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IsEmpty
ms.assetid: 493cbb8f-e069-14a9-a827-7f7631eb3a09
ms.date: 06/12/2019
localization_priority: Normal
---


# PictureFormat.IsEmpty property (Publisher)

Returns an **[MsoTriState](office.msotristate.md)** constant that represents whether the specified shape is an empty picture frame. Read-only.


## Syntax

_expression_.**IsEmpty**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

MsoTriState


## Remarks

The **IsEmpty** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The specified shape is not an empty picture frame.|
| **msoTriStateMixed**|Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The specified shape is an empty picture frame.|

## Example

The following example tests each picture in the active publication, and if it is not an empty picture frame, prints selected image properties for the picture.

```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 
 Debug.Print "File Name: " & .Filename 
 Debug.Print "Horizontal Scaling: " & .HorizontalScale & "%" 
 Debug.Print "Vertical Scaling: " & .VerticalScale & "%" 
 Debug.Print "File size in publication: " & .FileSize & " bytes" 
 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 
 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]