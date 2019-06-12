---
title: PictureFormat.OriginalIsTrueColor property (Publisher)
keywords: vbapb10.chm3604775
f1_keywords:
- vbapb10.chm3604775
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalIsTrueColor
ms.assetid: 837109d4-3479-2500-a1fa-b4c00e0f8672
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.OriginalIsTrueColor property (Publisher)

Returns an **[MsoTriState](office.msotristate.md)** constant indicating whether the specified linked picture or OLE object contains color data of 24 bits per channel or greater. Read-only.


## Syntax

_expression_.**OriginalIsTrueColor**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

MsoTriState


## Remarks

This property only applies to linked pictures or OLE objects. It returns "Permission Denied" for shapes representing embedded or pasted pictures and OLE objects.

Use either of the following properties to determine whether a shape represents a linked picture:

- The **[Type](Publisher.Shape.Type.md)** property of the **Shape** object   
- The **[IsLinked](Publisher.PictureFormat.IsLinked.md)** property of the **PictureFormat** object

The **OriginalIsTrueColor** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|The specified linked picture does not contain color data of 24 bits per channel or greater.|
| **msoTriStateMixed**|Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The specified linked picture contains color data of 24 bits per channel or greater.|

## Example

The following example returns a list of pictures in the active document that are TrueColor. If a picture is linked, and the linked picture is also TrueColor, that information is also returned.

```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture is TrueColor" 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoTrue Then 
 Debug.Print "The linked picture is also TrueColor." 
 End If 
 End If 
 End If 
 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]