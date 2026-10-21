---
title: Queries.FastCombine property (Excel)
keywords: vbaxl10.chm976078
f1_keywords:
- vbaxl10.chm976078
ms.assetid: 6d34ab2f-5dd4-6dd9-74c0-b49c600db45b
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# Queries.FastCombine property (Excel)

**True** if Power Query ignores privacy levels when it combines data from different data sources in the workbook (the fast combine feature). **False** if Power Query combines data according to the privacy level settings of each data source (see [Data Privacy Firewall](/power-query/data-privacy-firewall)). Read/write **Boolean**.


## Syntax

_expression_.**FastCombine**

_expression_ A variable that represents a **[Queries](excel.queries.md)** object.


## Remarks

The default value for a new workbook is **False**.

This property corresponds to the **Privacy** options for **Current Workbook** in the **Query Options** dialog box (on the **Data** tab, choose **Get Data** > **Query Options**). **True** corresponds to **Ignore the Privacy levels and potentially improve performance**, and **False** corresponds to **Combine data according to your Privacy Level settings for each source**.

This property takes effect only when the **Privacy** option in the **Global** section of the **Query Options** dialog box is set to **Combine data according to each file's Privacy Level settings**. If that option is set to **Always combine data according to your Privacy Level settings for each source** or **Always Ignore the Privacy levels and potentially improve performance**, the global option is used instead of this property.

The setting is stored in the workbook. If you set **FastCombine** and then save the workbook, the value is kept when the workbook is opened again. If you close the workbook without saving it, the change is discarded.

When **FastCombine** is **False**, queries that combine data from more than one data source can take longer to refresh or can fail with a `Formula.Firewall` error, depending on how the queries are structured. Setting **FastCombine** to **True** can avoid both, but data from one source can then be sent to another source. Set it to **True** only when all the data sources in the workbook can safely share data with each other.

For silent refresh operations, use the **FastCombine** property in conjunction with the **[Application.DisplayAlerts](Excel.Application.DisplayAlerts.md)** property set to **False**. 


## Example

This example ignores privacy levels for the active workbook and then refreshes all its queries. The data sources in this workbook are local files that the user owns.

```vb
Sub RefreshLocalQueries()
    ' This setting is saved with the workbook if the workbook is saved.
    ActiveWorkbook.Queries.FastCombine = True
    ActiveWorkbook.RefreshAll
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]