---
title: TableObject object (Excel)
keywords: vbaxl10.chm915072
f1_keywords:
- vbaxl10.chm915072
ms.prod: excel
ms.assetid: afc981f4-155b-085a-3c17-c8d46c4d7037
ms.date: 04/02/2019
localization_priority: Normal
---


# TableObject object (Excel)

Represents a worksheet table built from data returned from a PowerPivot model.


## Example

The following sample code creates a PowerPivot query table by connecting to a data source.

```vb
Sub CreateTable()
Dim objWBConnection As WorkbookConnection
Dim objWorksheet As Worksheet
Dim objTable As TableObject   'This is the new Table object

Set objWorksheet = ActiveWorkbook.Worksheets("Sheet1")

'Create a WorkbookConnection to the external data source first.
Set objWBConnection = ActiveWorkbook.Connections.Add2( _
        "Cubes3 AdventureWorksDW DimEmployee1", "", Array( _
        "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=AdventureWorksDW;Data Source=MyServer;Use " _
        , _
        "Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MYWORKSTATION;Use Encryption for Data=False;Tag with co" _
        , "lumn collation when possible=False"), Array( _
        """AdventureWorksDW"".""dbo"".""DimEmployee"""), 3, True)

'Create a new table connected to the model.
Set objTable = objWorksheet.ListObjects.Add(SourceType:=xlSrcModel, Source:=objWBConnection, Destination:=Range("$A$1")).TableObject

objTable.Refresh

End Sub

```


## Methods

- [Delete](Excel.tableobject.delete.md)
- [Refresh](Excel.tableobject.refresh.md)

## Properties

- [AdjustColumnWidth](Excel.tableobject.adjustcolumnwidth.md)
- [Application](Excel.tableobject.application.md)
- [Creator](Excel.tableobject.creator.md)
- [Destination](Excel.tableobject.destination.md)
- [EnableEditing](Excel.tableobject.enableediting.md)
- [EnableRefresh](Excel.tableobject.enablerefresh.md)
- [FetchedRowOverflow](Excel.tableobject.fetchedrowoverflow.md)
- [ListObject](Excel.tableobject.listobject.md)
- [Parent](Excel.tableobject.parent.md)
- [PreserveColumnInfo](Excel.tableobject.preservecolumninfo.md)
- [PreserveFormatting](Excel.tableobject.preserveformatting.md)
- [RefreshStyle](Excel.tableobject.refreshstyle.md)
- [ResultRange](Excel.tableobject.resultrange.md)
- [RowNumbers](Excel.tableobject.rownumbers.md)
- [WorkbookConnection](Excel.tableobject.workbookconnection.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
