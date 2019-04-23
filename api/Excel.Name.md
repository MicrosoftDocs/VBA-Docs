---
title: Name object (Excel)
keywords: vbaxl10.chm489072
f1_keywords:
- vbaxl10.chm489072
ms.prod: excel
api_name:
- Excel.Name
ms.assetid: cfedb297-ac0d-dff0-99c7-6927cc5f31ed
ms.date: 03/30/2019
localization_priority: Normal
---


# Name object (Excel)

Represents a defined name for a range of cells. Names can be either built-in names—such as Database, Print_Area, and Auto_Open—or custom names.


## Remarks

### Application, Workbook, and Worksheet objects

The **Name** object is a member of the **[Names](Excel.Names.md)** collection for the **[Application](Excel.Application(object).md)**, **[Workbook](Excel.Workbook.md)**, and **[Worksheet](Excel.Worksheet.md)** objects. Use **[Names](Excel.Workbook.Names.md)** (_index_), where _index_ is the name index number or defined name, to return a single **Name** object.

The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive.

### Range objects

Although a **[Range](Excel.Range(object).md)** object can have more than one name, there's no **Names** collection for the **Range** object. Use **[Name](Excel.Range.Name.md)** with a **Range** object to return the first name from the list of names (sorted alphabetically) assigned to the range. 

## Example

The following example displays the cell reference for the first name in the application collection.

```vb
MsgBox Names(1).RefersTo
```

<br/>

The following example deletes the name "mySortRange" from the active workbook.

```vb
ActiveWorkbook.Names("mySortRange").Delete
```

<br/>

Use the **Name** property to return or set the text of the name itself. The following example changes the name of the first **Name** object in the active workbook.

```vb
Names(1).Name = "stock_values"
```

<br/>

The following example sets the **[Visible](Excel.Worksheet.Visible.md)** property for the first name assigned to cells A1:B1 on worksheet one.

```vb
Worksheets(1).Range("a1:b1").Name.Visible = False
```


## Methods

- [Delete](Excel.Name.Delete.md)

## Properties

- [Application](Excel.Name.Application.md)
- [Category](Excel.Name.Category.md)
- [CategoryLocal](Excel.Name.CategoryLocal.md)
- [Comment](Excel.Name.Comment.md)
- [Creator](Excel.Name.Creator.md)
- [Index](Excel.Name.Index.md)
- [MacroType](Excel.Name.MacroType.md)
- [Name](Excel.Name.Name.md)
- [NameLocal](Excel.Name.NameLocal.md)
- [Parent](Excel.Name.Parent.md)
- [RefersTo](Excel.Name.RefersTo.md)
- [RefersToLocal](Excel.Name.RefersToLocal.md)
- [RefersToR1C1](Excel.Name.RefersToR1C1.md)
- [RefersToR1C1Local](Excel.Name.RefersToR1C1Local.md)
- [RefersToRange](Excel.Name.RefersToRange.md)
- [ShortcutKey](Excel.Name.ShortcutKey.md)
- [ValidWorkbookParameter](Excel.Name.ValidWorkbookParameter.md)
- [Value](Excel.Name.Value.md)
- [Visible](Excel.Name.Visible.md)
- [WorkbookParameter](Excel.Name.WorkbookParameter.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]