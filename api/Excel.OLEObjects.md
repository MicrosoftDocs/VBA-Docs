---
title: OLEObjects object (Excel)
keywords: vbaxl10.chm418072
f1_keywords:
- vbaxl10.chm418072
ms.prod: excel
api_name:
- Excel.OLEObjects
ms.assetid: e3fcf4bd-7c96-ecb3-dc04-551f7f7348f9
ms.date: 03/30/2019
localization_priority: Normal
---


# OLEObjects object (Excel)

A collection of all the **[OLEObject](Excel.OLEObject.md)** objects on the specified worksheet.


## Remarks

Each **OLEObject** object represents an ActiveX control or a linked or embedded OLE object.

An ActiveX control on a sheet has two names: the name of the shape that contains the control, which you can see in the **Name** box when you view the sheet, and the code name for the control, which you can see in the cell to the right of **(Name)** in the Properties window. 

When you first add a control to a sheet, the shape name and code name match. However, if you change either the shape name or code name, the other is not automatically changed to match.


## Example

Use the **[OLEObjects](Excel.Worksheet.OLEObjects.md)** method of the **Worksheet** object to return the **OLEObjects** collection. 

The following example hides all the OLE objects on worksheet one.

```vb
Worksheets(1).OLEObjects.Visible = False
```

<br/>

Use the **Add** method to create a new OLE object and add it to the **OLEObjects** collection. The following example creates a new OLE object representing the bitmap file Arcade.bmp and adds it to worksheet one.

```vb
Worksheets(1).OLEObjects.Add FileName:="arcade.gif"
```

<br/>

The following example creates a new ActiveX control (a list box) and adds it to worksheet one.

```vb
Worksheets(1).OLEObjects.Add ClassType:="Forms.ListBox.1"
```

<br/>

You use the code name of a control in the names of its event procedures. However, when you return a control from the **[Shapes](Excel.Shapes.md)** or **OLEObjects** collection for a sheet, you must use the shape name, not the code name, to refer to the control by name. For example, assume that you add a check box to a sheet and that both the default shape name and the default code name are CheckBox1. If you then change the control code name by typing chkFinished next to **(Name)** in the Properties window, you must use chkFinished in event procedures names, but you still have to use CheckBox1 to return the control from the **Shapes** or **OLEObject** collection, as shown in the following example.

```vb
Private Sub chkFinished_Click() 
 ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 
End Sub
```


## Methods

- [Add](Excel.OLEObjects.Add.md)
- [BringToFront](Excel.OLEObjects.BringToFront.md)
- [Copy](Excel.OLEObjects.Copy.md)
- [CopyPicture](Excel.OLEObjects.CopyPicture.md)
- [Cut](Excel.OLEObjects.Cut.md)
- [Delete](Excel.OLEObjects.Delete.md)
- [Duplicate](Excel.OLEObjects.Duplicate.md)
- [Item](Excel.OLEObjects.Item.md)
- [Select](Excel.OLEObjects.Select.md)
- [SendToBack](Excel.OLEObjects.SendToBack.md)

## Properties

- [Application](Excel.OLEObjects.Application.md)
- [AutoLoad](Excel.OLEObjects.AutoLoad.md)
- [Border](Excel.OLEObjects.Border.md)
- [Count](Excel.OLEObjects.Count.md)
- [Creator](Excel.OLEObjects.Creator.md)
- [Enabled](Excel.OLEObjects.Enabled.md)
- [Height](Excel.OLEObjects.Height.md)
- [Interior](Excel.OLEObjects.Interior.md)
- [Left](Excel.OLEObjects.Left.md)
- [Locked](Excel.OLEObjects.Locked.md)
- [Parent](Excel.OLEObjects.Parent.md)
- [Placement](Excel.OLEObjects.Placement.md)
- [PrintObject](Excel.OLEObjects.PrintObject.md)
- [Shadow](Excel.OLEObjects.Shadow.md)
- [ShapeRange](Excel.OLEObjects.ShapeRange.md)
- [SourceName](Excel.OLEObjects.SourceName.md)
- [Top](Excel.OLEObjects.Top.md)
- [Visible](Excel.OLEObjects.Visible.md)
- [Width](Excel.OLEObjects.Width.md)
- [ZOrder](Excel.OLEObjects.ZOrder.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]