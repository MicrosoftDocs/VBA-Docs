---
title: Chart.SeriesNameLevel property (Excel)
keywords: vbaxl10.chm149196
f1_keywords:
- vbaxl10.chm149196
ms.assetid: 17ada484-943e-502f-a499-077d1e53e6c1
ms.date: 04/16/2019
ms.localizationpriority: medium
---


# Chart.SeriesNameLevel property (Excel)

Returns an **[XlSeriesNameLevel](Excel.xlseriesnamelevel.md)** constant referring to the level of where the series names are being sourced from. Read/write **Integer**.


## Syntax

_expression_.**SeriesNameLevel**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

If there is a hierarchy, 0 refers to the parent level, 1 refers to its children, and so on. So, 0 equals the first level, 1 equals the second level, 2 equals the third level, and so on.


## Property value

**XLSERIESNAMELEVEL**


## Example

The following sample code uses the **SeriesNameLevel** property to set the chart series names from previously created columns.

```vb
    Sheets(1).Range("C1:E1").Value2 = "Sample_Row1"
    Sheets(1).Range("C2:E2").Value2 = "Sample_Row2"
    Sheets(1).Range("A3:A5").Value2 = "Sample_ColA"
    Sheets(1).Range("B3:B5").Value2 = "Sample_ColB"
    Sheets(1).Range("C3:E5").Formula = "=row()"
    Dim crt As Chart
    Set crt = Sheets(1).ChartObjects.Add(0, 0, 500, 200).Chart
    crt.SetSourceData Sheets(1).Range("A1:E5")
    ' Set the series names to only use column B
    crt.SeriesNameLevel = 1
    ' Use columns A and B for the series names
    crt.SeriesNameLevel = xlSeriesNameLevelAll
    ' Use row 1 for the category labels
    crt.CategoryLabelLevel = 0
    ' Use rows 1 and 2 for the category labels
    crt.CategoryLabelLevel = xlCategoryLabelLevelAll
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]