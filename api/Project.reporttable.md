---
title: ReportTable object (Project)
ms.prod: project-server
ms.assetid: db9846c7-fd53-ae5a-7a43-35dfc60f4fe4
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportTable object (Project)
Represents a shape in the form of a table in a Project report.
 

## Remarks


> [!NOTE] 
> Macro recording for the **ReportTable** object is not implemented. That is, when you record a macro in Project and manually add a report table or edit table elements, the steps for adding and manipulating the report table are not recorded.
 

The **ReportTable** object is a kind of **Shape** object; it is not related to the **Table** object. Project has limited VBA support for report tables; to specify the table fields, you manually use the **Field List** task pane (see Figure 1). To show or hide the **Field List** task pane, choose the **Table Data** command in the **DESIGN** tab under **TABLE TOOLS** on the ribbon. To specify the table layout or design properties, you can use the **DESIGN** tab and the **LAYOUT** tab on the ribbon.
 

 
You can update the data query associated with a report table, by using the **[UpdateTableData](Project.reporttable.updatetabledata.md)** method. To get the text in a table cell, use the **[GetCellText](Project.reporttable.getcelltext.md)** method.
 

 
To programmatically create a **ReportTable**, use the **[Shapes.AddTable](Project.shapes.addtable.md)** method. To return a **ReportTable** object, use `Shapes(Index).Table`, where  _Index_ is the name or the index number of a shape.
 

 

## Example

The **TestReportTable** macro creates a report named Table Tests, and then creates a **ReportTable** object.
 

 

```vb
Sub TestReportTable()
    Dim theReport As Report
    Dim theShape As Shape
    Dim theReportTable As ReportTable
    Dim reportName As String
    Dim tableName As String
    Dim rows As Integer, columns As Integer, left As Integer, _
        top As Integer, width As Integer, height As Integer    
    rows = 3
    columns = 4
    left = 20
    top = 20
    width = 200
    height = 100
    
    reportName = "Table Tests"
    tableName = "Basic Project Data Table"
    
    Set theReport = ActiveProject.Reports.Add(reportName)
    
    ' Project ignores the NumRows and NumColumns parameters when creating a ReportTable.
    Set theShape = theReport.Shapes.AddTable( _
        rows, columns, left, top, width, height)
    
    theShape.Name = tableName
    
    Set theReportTable = theShape.Table
    
    With theReportTable
        Debug.Print "Rows: " & .RowsCount
        Debug.Print "Columns: " & .ColumnsCount
        Debug.Print "Table contents:" & vbCrLf & .GetCellText(1, 1)
    End With
End Sub
```

In Figure 1, the top  **ReportTable** object in the Table Tests report is created by the **TestReportTable** macro. When you first create the table, it has one row and one column; the _NumRows_ and _NumColumns_ parameters of the **AddTable** method have no effect. The number of rows and columns in the table is updated when you manually add fields to the table from the **Field List** task pane, or when you use the [UpdateTableData](Project.reporttable.updatetabledata.md) method. You can filter the fields to limit the number of rows. The **TestReportTable** macro writes the following in the Immediate window of the VBE:
 

 



```vb
Rows: 1
Columns: 1
Table contents:
Use the Table Data taskpane to build a table
```

The bottom  **ReportTable** object in Figure 1 is the default report table that Project creates when you choose **Table** on the **DESIGN** tab under **REPORT TOOLS**. It shows the project name, start date, finish date, and percent complete of the project summary task (task ID = 0).
 

 

**Figure 1. The ReportTable object requires manual editing to add fields and change formatting**

 
![The ReportTable object requires manual editing](../images/pj15_VBA_ReportTableObject.gif)To delete a **ReportTable** object, use the **[Shape.Delete](Project.shape.delete.md)** method, as in the following macro:
 

 



```vb
Sub DeleteTheReportTable()
    Dim theReport As Report
    Dim theShape As Shape
    Dim reportName As String
    Dim tableName As String
    
    reportName = "Table Tests"
    tableName = "Basic Project Data Table"
    
    Set theReport = ActiveProject.Reports(reportName)
    Set theShape = theReport.Shapes(tableName)
    
    theShape.Delete
End Sub
```

To delete the entire report, change to another view, as in the following macro:
 

 



```vb
Sub DeleteTheReport()
    Dim i As Integer
    Dim reportName As String
    
    reportName = "Table Tests"
    
    ' To delete the active report, change to another view.
    ViewApplyEx Name:="&Gantt Chart"
    
    ActiveProject.Reports(reportName).Delete
End Sub
```


## Methods



|Name|
|:-----|
|[GetCellText](Project.reporttable.getcelltext.md)|
|[UpdateTableData](Project.reporttable.updatetabledata.md)|

## Properties



|Name|
|:-----|
|[ColumnsCount](Project.reporttable.columnscount.md)|
|[RowsCount](Project.reporttable.rowscount.md)|

## See also


 
[Report Object](Project.report.md)
 
[Shape Object](Project.shape.md)
 
[Chart Object](Project.chart.md)
 
[Chart.DataTable Property](Project.chart.datatable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]