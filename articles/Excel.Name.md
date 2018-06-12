---
title: Name Object (Excel)
keywords: vbaxl10.chm489072
f1_keywords:
- vbaxl10.chm489072
ms.prod: excel
api_name:
- Excel.Name
ms.assetid: cfedb297-ac0d-dff0-99c7-6927cc5f31ed
ms.date: 06/08/2017
---


# Name Object (Excel)

Represents a defined name for a range of cells. Names can be either built-in names — such as Database, Print_Area, and Auto_Open — or custom names.


## Remarks

 **Application, Workbook, and Worksheet Objects**

The  **Name** object is a member of the **[Names](Excel.Names.md)** collection for the **[Application](Excel.Application(objec).md)**, **[Workbook](Excel.Workbook.md)**, and **[Worksheet](Excel.Worksheet.md)** objects. Use **[Names](Excel.Workbook.Names.md)** ( _index_ ), where _index_ is the name index number or defined name, to return a single **Name** object.

The index number indicates the position of the name within the collection. Names are placed in alphabetic order, from a to z, and are not case-sensitive.

 **Range Objects**

Although a  **[Range](Excel.Range(objec).md)** object can have more than one name, there's no **Names** collection for the **Range** object. Use **[Name](Excel.Range.Name.md)** with a **Range** object to return the first name from the list of names (sorted alphabetically) assigned to the range. The following example sets the **[Visible](Excel.Worksheet.Visible.md)** property for the first name assigned to cells A1:B1 on worksheet one.


## Example

The following example displays the cell reference for the first name in the application collection.


```
MsgBox Names(1).RefersTo
```

The following example deletes the name "mySortRange" from the active workbook.




```
ActiveWorkbook.Names("mySortRange").Delete
```

Use the  **Name** property to return or set the text of the name itself. The following example changes the name of the first **Name** object in the active workbook.




```
Names(1).Name = "stock_values"
```

The following example sets the  **Visible** property for the first name assigned to cells A1:B1 on worksheet one.




```
Worksheets(1).Range("a1:b1").Name.Visible = False
```


## Methods



|**Name**|
|:-----|
|[Delete](Excel.Name.Delete.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Name.Application.md)|
|[Category](Excel.Name.Category.md)|
|[CategoryLocal](Excel.Name.CategoryLocal.md)|
|[Comment](Excel.Name.Comment.md)|
|[Creator](Excel.Name.Creator.md)|
|[Index](Excel.Name.Index.md)|
|[MacroType](Excel.Name.MacroType.md)|
|[Name](Excel.Name.Name.md)|
|[NameLocal](Excel.Name.NameLocal.md)|
|[Parent](Excel.Name.Parent.md)|
|[RefersTo](Excel.Name.RefersTo.md)|
|[RefersToLocal](Excel.Name.RefersToLocal.md)|
|[RefersToR1C1](Excel.Name.RefersToR1C1.md)|
|[RefersToR1C1Local](Excel.Name.RefersToR1C1Local.md)|
|[RefersToRange](Excel.Name.RefersToRange.md)|
|[ShortcutKey](Excel.Name.ShortcutKey.md)|
|[ValidWorkbookParameter](Excel.Name.ValidWorkbookParameter.md)|
|[Value](Excel.Name.Value.md)|
|[Visible](Excel.Name.Visible.md)|
|[WorkbookParameter](Excel.Name.WorkbookParameter.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
