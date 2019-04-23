---
title: OLEObject object (Excel)
keywords: vbaxl10.chm414072
f1_keywords:
- vbaxl10.chm414072
ms.prod: excel
api_name:
- Excel.OLEObject
ms.assetid: bc3ef12d-1531-6c21-71ab-3df6bb851f3b
ms.date: 03/30/2019
localization_priority: Normal
---


# OLEObject object (Excel)

Represents an ActiveX control or a linked or embedded OLE object on a worksheet.


## Remarks

The **OLEObject** object is a member of the **[OLEObjects](Excel.OLEObjects.md)** collection. The **OLEObjects** collection contains all the OLE objects on a single worksheet.


## Example

Use **[OLEObjects](Excel.Worksheet.OLEObjects.md)** (_index_), where _index_ is the name or number of the object, to return an **OLEObject** object. 

The following example deletes OLE object one on Sheet1.

```vb
Worksheets("sheet1").OLEObjects(1).Delete
```

<br/>

The following example deletes the OLE object named ListBox1.

```vb
Worksheets("sheet1").OLEObjects("ListBox1").Delete
```

<br/>

The properties and methods of the **OLEObject** object are duplicated on each ActiveX control on a worksheet. This enables Visual Basic code to gain access to these properties by using the control's name. The following example selects the check box control named **MyCheckBox**, aligns it with the active cell, and then activates the control.

```vb
With MyCheckBox 
 .Value = True 
 .Top = ActiveCell.Top 
 .Activate 
End With
```


## Events

- [GotFocus](Excel.OLEObject.GotFocus.md)
- [LostFocus](Excel.OLEObject.LostFocus.md)

## Methods

- [Activate](Excel.OLEObject.Activate.md)
- [BringToFront](Excel.OLEObject.BringToFront.md)
- [Copy](Excel.OLEObject.Copy.md)
- [CopyPicture](Excel.OLEObject.CopyPicture.md)
- [Cut](Excel.OLEObject.Cut.md)
- [Delete](Excel.OLEObject.Delete.md)
- [Duplicate](Excel.OLEObject.Duplicate.md)
- [Select](Excel.OLEObject.Select.md)
- [SendToBack](Excel.OLEObject.SendToBack.md)
- [Update](Excel.OLEObject.Update.md)
- [Verb](Excel.OLEObject.Verb.md)

## Properties

- [Application](Excel.OLEObject.Application.md)
- [AutoLoad](Excel.OLEObject.AutoLoad.md)
- [AutoUpdate](Excel.OLEObject.AutoUpdate.md)
- [Border](Excel.OLEObject.Border.md)
- [BottomRightCell](Excel.OLEObject.BottomRightCell.md)
- [Creator](Excel.OLEObject.Creator.md)
- [Enabled](Excel.OLEObject.Enabled.md)
- [Height](Excel.OLEObject.Height.md)
- [Index](Excel.OLEObject.Index.md)
- [Interior](Excel.OLEObject.Interior.md)
- [Left](Excel.OLEObject.Left.md)
- [LinkedCell](Excel.OLEObject.LinkedCell.md)
- [ListFillRange](Excel.OLEObject.ListFillRange.md)
- [Locked](Excel.OLEObject.Locked.md)
- [Name](Excel.OLEObject.Name.md)
- [Object](Excel.OLEObject.Object.md)
- [OLEType](Excel.OLEObject.OLEType.md)
- [Parent](Excel.OLEObject.Parent.md)
- [Placement](Excel.OLEObject.Placement.md)
- [PrintObject](Excel.OLEObject.PrintObject.md)
- [progID](Excel.OLEObject.progID.md)
- [Shadow](Excel.OLEObject.Shadow.md)
- [ShapeRange](Excel.OLEObject.ShapeRange.md)
- [SourceName](Excel.OLEObject.SourceName.md)
- [Top](Excel.OLEObject.Top.md)
- [TopLeftCell](Excel.OLEObject.TopLeftCell.md)
- [Visible](Excel.OLEObject.Visible.md)
- [Width](Excel.OLEObject.Width.md)
- [ZOrder](Excel.OLEObject.ZOrder.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]