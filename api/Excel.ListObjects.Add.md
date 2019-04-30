---
title: ListObjects.Add method (Excel)
keywords: vbaxl10.chm732078
f1_keywords:
- vbaxl10.chm732078
ms.prod: excel
api_name:
- Excel.ListObjects.Add
ms.assetid: 764dafed-d4e3-82b9-df8c-68a358319491
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObjects.Add method (Excel)

Creates a new list object.


## Syntax

_expression_.**Add** (_SourceType_, _Source_, _LinkSource_, _XlListObjectHasHeaders_, _Destination_, _TableStyleName_)

_expression_ A variable that represents a **[ListObjects](Excel.ListObjects.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SourceType_|Optional|**[XlListObjectSourceType](Excel.XlListObjectSourceType.md)**|Indicates the kind of source for the query. |
| _Source_|Optional|**Variant**|When _SourceType_ = **xlSrcRange**: A **[Range](Excel.Range(object).md)** object representing the data source. If omitted, the _Source_ will default to the range returned by list range detection code.<br/><br/>When _SourceType_ = **xlSrcExternal**: An array of **String** values specifying a connection to the source, containing the following elements:<ul><li>0 - URL to SharePoint site</li><li>1 - ListName</li><li>2 - ViewGUID</li></ul>|
| _LinkSource_|Optional|**Boolean**| Indicates whether an external data source is to be linked to the **ListObject** object. If _SourceType_ is **xlSrcExternal**, the default is **True**. Invalid if _SourceType_ is **xlSrcRange**, and will return an error if not omitted.|
| _XlListObjectHasHeaders_|Optional|**Variant**|An **[XlYesNoGuess](Excel.XlYesNoGuess.md)** constant that indicates whether the data being imported has column labels. If the _Source_ does not contain headers, Excel will automatically generate headers. Default value: **xlGuess**.|
| _Destination_|Optional|**Variant**|A **Range** object specifying a single-cell reference as the destination for the top-left corner of the new list object. If the **Range** object refers to more than one cell, an error is generated.<br/><br/>The _Destination_ argument must be specified when _SourceType_ is set to **xlSrcExternal**. The _Destination_ argument is ignored if _SourceType_ is set to **xlSrcRange**.<br/><br/>The destination range must be on the worksheet that contains the **ListObjects** collection specified by _expression_. New columns will be inserted at the _Destination_ to fit the new list. Therefore, existing data will not be overwritten.|
| _TableStyleName_|Optional|**String**| The name of a **[TableStyle](Excel.TableStyle.md)**; for example "TableStyleLight1". |

## Return value

A **[ListObject](Excel.ListObject.md)** object that represents the new list object.


## Remarks

When the list has headers, the first row of cells will be converted to **Text**, if not already set to text. The conversion will be based on the visible text for the cell. This means that if there is a date value with a **Date** format that changes with locale, the conversion to a list might produce different results depending on the current system locale. Moreover, if there are two cells in the header row that have the same visible text, an incremental **Integer** will be appended to make each column header unique.


## Example

The following example adds a new **ListObject** object based on data from a Microsoft SharePoint Foundation site to the default **ListObjects** collection and places the list in cell A1 in the first worksheet of the workbook.

> [!NOTE] 
> The following code example assumes that you will substitute a valid server name and the list guid in the variables  `strServerName` and `strListGUID`. Additionally, the server name must be followed by `"/_vti_bin" (strListName)` or the sample will not work.


```vb
Set objListObject = ActiveWorkbook.Worksheets(1).ListObjects.Add(SourceType:= xlSrcExternal, _ 
Source:= Array(strServerName, strListName, strListGUID), LinkSource:=True, _ 
XlListObjectHasHeaders:=xlGuess, Destination:=Range("A1")), 
TableStyleName:=xlGuess, Destination:=Range("A10")) 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
