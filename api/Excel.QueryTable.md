---
title: QueryTable object (Excel)
keywords: vbaxl10.chm517072
f1_keywords:
- vbaxl10.chm517072
ms.prod: excel
api_name:
- Excel.QueryTable
ms.assetid: 505b84ea-64b3-b4fe-741a-de6884eb69eb
ms.date: 04/02/2019
localization_priority: Normal
---


# QueryTable object (Excel)

Represents a worksheet table built from data returned from an external data source, such as a SQL server or a Microsoft Access database.


## Remarks

The **QueryTable** object is a member of the **[QueryTables](Excel.QueryTables.md)** collection.


## Example

Use **[QueryTables](Excel.Worksheet.QueryTables.md)** (_index_), where _index_ is the index number of the query table, to return a single **QueryTable** object. 

The following example sets query table one so that formulas to the right of it are automatically updated whenever it's refreshed.

```vb
Sheets("sheet1").QueryTables(1).FillAdjacentFormulas = True
```

## Events

- [AfterRefresh](Excel.QueryTable.AfterRefresh.md)
- [BeforeRefresh](Excel.QueryTable.BeforeRefresh.md)

## Methods

- [CancelRefresh](Excel.QueryTable.CancelRefresh.md)
- [Delete](Excel.QueryTable.Delete.md)
- [Refresh](Excel.QueryTable.Refresh.md)
- [ResetTimer](Excel.QueryTable.ResetTimer.md)
- [SaveAsODC](Excel.QueryTable.SaveAsODC.md)

## Properties

- [AdjustColumnWidth](Excel.QueryTable.AdjustColumnWidth.md)
- [Application](Excel.QueryTable.Application.md)
- [BackgroundQuery](Excel.QueryTable.BackgroundQuery.md)
- [CommandText](Excel.QueryTable.CommandText.md)
- [CommandType](Excel.QueryTable.CommandType.md)
- [Connection](Excel.QueryTable.Connection.md)
- [Creator](Excel.QueryTable.Creator.md)
- [Destination](Excel.QueryTable.Destination.md)
- [EditWebPage](Excel.QueryTable.EditWebPage.md)
- [EnableEditing](Excel.QueryTable.EnableEditing.md)
- [EnableRefresh](Excel.QueryTable.EnableRefresh.md)
- [FetchedRowOverflow](Excel.QueryTable.FetchedRowOverflow.md)
- [FieldNames](Excel.QueryTable.FieldNames.md)
- [FillAdjacentFormulas](Excel.QueryTable.FillAdjacentFormulas.md)
- [ListObject](Excel.QueryTable.ListObject.md)
- [MaintainConnection](Excel.QueryTable.MaintainConnection.md)
- [Name](Excel.QueryTable.Name.md)
- [Parameters](Excel.QueryTable.Parameters.md)
- [Parent](Excel.QueryTable.Parent.md)
- [PostText](Excel.QueryTable.PostText.md)
- [PreserveColumnInfo](Excel.QueryTable.PreserveColumnInfo.md)
- [PreserveFormatting](Excel.QueryTable.PreserveFormatting.md)
- [QueryType](Excel.QueryTable.QueryType.md)
- [Recordset](Excel.QueryTable.Recordset.md)
- [Refreshing](Excel.QueryTable.Refreshing.md)
- [RefreshOnFileOpen](Excel.QueryTable.RefreshOnFileOpen.md)
- [RefreshPeriod](Excel.QueryTable.RefreshPeriod.md)
- [RefreshStyle](Excel.QueryTable.RefreshStyle.md)
- [ResultRange](Excel.QueryTable.ResultRange.md)
- [RobustConnect](Excel.QueryTable.RobustConnect.md)
- [RowNumbers](Excel.QueryTable.RowNumbers.md)
- [SaveData](Excel.QueryTable.SaveData.md)
- [SavePassword](Excel.QueryTable.SavePassword.md)
- [Sort](Excel.QueryTable.Sort.md)
- [SourceConnectionFile](Excel.QueryTable.SourceConnectionFile.md)
- [SourceDataFile](Excel.QueryTable.SourceDataFile.md)
- [TextFileColumnDataTypes](Excel.QueryTable.TextFileColumnDataTypes.md)
- [TextFileCommaDelimiter](Excel.QueryTable.TextFileCommaDelimiter.md)
- [TextFileConsecutiveDelimiter](Excel.QueryTable.TextFileConsecutiveDelimiter.md)
- [TextFileDecimalSeparator](Excel.QueryTable.TextFileDecimalSeparator.md)
- [TextFileFixedColumnWidths](Excel.QueryTable.TextFileFixedColumnWidths.md)
- [TextFileOtherDelimiter](Excel.QueryTable.TextFileOtherDelimiter.md)
- [TextFileParseType](Excel.QueryTable.TextFileParseType.md)
- [TextFilePlatform](Excel.QueryTable.TextFilePlatform.md)
- [TextFilePromptOnRefresh](Excel.QueryTable.TextFilePromptOnRefresh.md)
- [TextFileSemicolonDelimiter](Excel.QueryTable.TextFileSemicolonDelimiter.md)
- [TextFileSpaceDelimiter](Excel.QueryTable.TextFileSpaceDelimiter.md)
- [TextFileStartRow](Excel.QueryTable.TextFileStartRow.md)
- [TextFileTabDelimiter](Excel.QueryTable.TextFileTabDelimiter.md)
- [TextFileTextQualifier](Excel.QueryTable.TextFileTextQualifier.md)
- [TextFileThousandsSeparator](Excel.QueryTable.TextFileThousandsSeparator.md)
- [TextFileTrailingMinusNumbers](Excel.QueryTable.TextFileTrailingMinusNumbers.md)
- [TextFileVisualLayout](Excel.QueryTable.TextFileVisualLayout.md)
- [WebConsecutiveDelimitersAsOne](Excel.QueryTable.WebConsecutiveDelimitersAsOne.md)
- [WebDisableDateRecognition](Excel.QueryTable.WebDisableDateRecognition.md)
- [WebDisableRedirections](Excel.QueryTable.WebDisableRedirections.md)
- [WebFormatting](Excel.QueryTable.WebFormatting.md)
- [WebPreFormattedTextToColumns](Excel.QueryTable.WebPreFormattedTextToColumns.md)
- [WebSelectionType](Excel.QueryTable.WebSelectionType.md)
- [WebSingleBlockTextImport](Excel.QueryTable.WebSingleBlockTextImport.md)
- [WebTables](Excel.QueryTable.WebTables.md)
- [WorkbookConnection](Excel.QueryTable.WorkbookConnection.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
