---
title: PivotCache Object (Excel)
keywords: vbaxl10.chm226072
f1_keywords:
- vbaxl10.chm226072
ms.prod: excel
api_name:
- Excel.PivotCache
ms.assetid: c3d84ef1-f9e6-b1bc-cbf0-3ba8dfe17439
ms.date: 06/08/2017
---


# PivotCache Object (Excel)

Represents the memory cache for a PivotTable report.


## Remarks

 The **PivotCache** object is a member of the **[PivotCaches](Excel.PivotCaches.md)** collection.


## Example

Use the  **[PivotCache](Excel.PivotTable.PivotCache.md)** method to return a **PivotCache** object for a PivotTable report (each report has only one cache). The following example causes the first PivotTable report on the first worksheet to refresh itself whenever its file is opened.


```
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True
```

Use  **[PivotCaches](Excel.Workbook.PivotCaches.md)** ( _index_ ), where _index_ is the PivotTable cache number, to return a single **PivotCache** object from the **PivotCaches** collection for a workbook. The following example refreshes cache one.




```
ActiveWorkbook.PivotCaches(1).Refresh
```


## Methods



|**Name**|
|:-----|
|[CreatePivotChart](Excel.pivotcache.createpivotchart.md)|
|[CreatePivotTable](Excel.PivotCache.CreatePivotTable.md)|
|[MakeConnection](Excel.PivotCache.MakeConnection.md)|
|[Refresh](Excel.PivotCache.Refresh.md)|
|[ResetTimer](Excel.PivotCache.ResetTimer.md)|
|[SaveAsODC](Excel.PivotCache.SaveAsODC.md)|

## Properties



|**Name**|
|:-----|
|[ADOConnection](Excel.PivotCache.ADOConnection.md)|
|[Application](Excel.PivotCache.Application.md)|
|[BackgroundQuery](Excel.PivotCache.BackgroundQuery.md)|
|[CommandText](Excel.PivotCache.CommandText.md)|
|[CommandType](Excel.PivotCache.CommandType.md)|
|[Connection](Excel.PivotCache.Connection.md)|
|[Creator](Excel.PivotCache.Creator.md)|
|[EnableRefresh](Excel.PivotCache.EnableRefresh.md)|
|[Index](Excel.PivotCache.Index.md)|
|[IsConnected](Excel.PivotCache.IsConnected.md)|
|[LocalConnection](Excel.PivotCache.LocalConnection.md)|
|[MaintainConnection](Excel.PivotCache.MaintainConnection.md)|
|[MemoryUsed](Excel.PivotCache.MemoryUsed.md)|
|[MissingItemsLimit](Excel.PivotCache.MissingItemsLimit.md)|
|[OLAP](Excel.PivotCache.OLAP.md)|
|[OptimizeCache](Excel.PivotCache.OptimizeCache.md)|
|[Parent](Excel.PivotCache.Parent.md)|
|[QueryType](Excel.PivotCache.QueryType.md)|
|[RecordCount](Excel.PivotCache.RecordCount.md)|
|[Recordset](Excel.PivotCache.Recordset.md)|
|[RefreshDate](Excel.PivotCache.RefreshDate.md)|
|[RefreshName](Excel.PivotCache.RefreshName.md)|
|[RefreshOnFileOpen](Excel.PivotCache.RefreshOnFileOpen.md)|
|[RefreshPeriod](Excel.PivotCache.RefreshPeriod.md)|
|[RobustConnect](Excel.PivotCache.RobustConnect.md)|
|[SavePassword](Excel.PivotCache.SavePassword.md)|
|[SourceConnectionFile](Excel.PivotCache.SourceConnectionFile.md)|
|[SourceData](Excel.PivotCache.SourceData.md)|
|[SourceDataFile](Excel.PivotCache.SourceDataFile.md)|
|[SourceType](Excel.PivotCache.SourceType.md)|
|[UpgradeOnRefresh](Excel.PivotCache.UpgradeOnRefresh.md)|
|[UseLocalConnection](Excel.PivotCache.UseLocalConnection.md)|
|[Version](Excel.PivotCache.Version.md)|
|[WorkbookConnection](Excel.PivotCache.WorkbookConnection.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
