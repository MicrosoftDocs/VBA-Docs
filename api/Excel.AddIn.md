---
title: AddIn object (Excel)
keywords: vbaxl10.chm184072
f1_keywords:
- vbaxl10.chm184072
ms.prod: excel
api_name:
- Excel.AddIn
ms.assetid: ad26800d-5342-fb4c-01f3-05b7eceb7ffd
ms.date: 03/29/2019
localization_priority: Normal
---


# AddIn object (Excel)

Represents a single add-in, either installed or not installed. 


## Remarks

The **AddIn** object is a member of the **[AddIns](Excel.AddIns.md)** collection. The **AddIns** collection contains a list of all the add-ins available to Microsoft Excel, regardless of whether they're installed. This list corresponds to the list of add-ins displayed in the **Add-Ins** dialog box.

## Example

Use **AddIns** (_index_), where _index_ is the add-in title or index number, to return a single **AddIn** object. The following example installs the Analysis Toolpak add-in.

```vb
AddIns("analysis toolpak").Installed = True
```

<br/>

Don't confuse the add-in title, which appears in the **Add-Ins** dialog box, with the add-in name, which is the file name of the add-in. You must spell the add-in title exactly as it's spelled in the **Add-Ins** dialog box, but the capitalization doesn't have to match.

The index number represents the position of the add-in in the **Add-ins available** box in the **Add-Ins** dialog box. The following example creates a list that contains specified properties of the available add-ins.

```vb
With Worksheets("sheet1") 
 .Rows(1).Font.Bold = True 
 .Range("a1:d1").Value = _ 
 Array("Name", "Full Name", "Title", "Installed") 
 For i = 1 To AddIns.Count 
 .Cells(i + 1, 1) = AddIns(i).Name 
 .Cells(i + 1, 2) = AddIns(i).FullName 
 .Cells(i + 1, 3) = AddIns(i).Title 
 .Cells(i + 1, 4) = AddIns(i).Installed 
 Next 
 .Range("a1").CurrentRegion.Columns.AutoFit 
End With
```

<br/>

The **[Add](Excel.AddIns.Add.md)** method adds an add-in to the list of available add-ins but doesn't install the add-in. Set the **Installed** property of the add-in to **True** to install the add-in. 

To install an add-in that doesn't appear in the list of available add-ins, you must first use the **Add** method and then set the **Installed** property. This can be done in a single step, as shown in the following example (note that you use the name of the add-in, not its title, with the **Add** method).

```vb
AddIns.Add("generic.xll").Installed = True
```

<br/>

Use **Workbooks** (_index_), where _index_ is the add-in file name (not title) to return a reference to the workbook corresponding to a loaded add-in. You must use the file name because loaded add-ins don't normally appear in the **Workbooks** collection. This example sets the _wb_ variable to the workbook for Myaddin.xla.

```vb
Set wb = Workbooks("myaddin.xla")
```

<br/>

The following example sets the _wb_ variable to the workbook for the Analysis Toolpak add-in.

```vb
Set wb = Workbooks(AddIns("analysis toolpak").Name)
```

<br/>

If the **Installed** property returns **True**, but the calls to functions in the add-in still fail, the add-in may not actually be loaded. This is because the **Addin** object represents the existence and installed state of the add-in but doesn't represent the actual contents of the add-in workbook.To guarantee that an installed add-in is loaded, you should open the add-in workbook. 

The following example opens the workbook for the add-in named "My Addin" if the add-in isn't already present in the **Workbooks** collection.

```vb
On Error Resume Next ' turn off error checking 
Set wbMyAddin = Workbooks(AddIns("My Addin").Name) 
lastError = Err 
On Error Goto 0 ' restore error checking 
If lastError <> 0 Then 
 ' the add-in workbook isn't currently open. Manually open it. 
 Set wbMyAddin = Workbooks.Open(AddIns("My Addin").FullName) 
End If
```


## Properties

- [Application](Excel.AddIn.Application.md)
- [CLSID](Excel.AddIn.CLSID.md)
- [Creator](Excel.AddIn.Creator.md)
- [FullName](Excel.AddIn.FullName.md)
- [Installed](Excel.AddIn.Installed.md)
- [IsOpen](Excel.AddIn.IsOpen.md)
- [Name](Excel.AddIn.Name.md)
- [Parent](Excel.AddIn.Parent.md)
- [Path](Excel.AddIn.Path.md)
- [progID](Excel.AddIn.progID.md)

## See also

- [Excel Object Model reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
