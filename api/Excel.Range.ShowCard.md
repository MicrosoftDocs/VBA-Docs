---
title: Range.ShowCard method (Excel)
keywords: vbaxl10.chm144258
f1_keywords:
- vbaxl10.chm144258
ms.prod: excel
api_name:
- Excel.Range.ShowCard
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.ShowCard method (Excel)

For a cell containing a Linked data type, such as [Stocks or Geography](https://support.office.com/article/stock-quotes-and-geographic-data-61a33056-9935-484f-8ac8-f1a89e210877), this method causes a card to appear that shows details about the cell (that is, the same card that the user can view by choosing the cell icon).

## Syntax

_expression_.**ShowCard**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.

## Remarks

For ranges of more than one cell, this method only attempts to show the card for the upper-left cell in the range. If that cell does not contain a Linked data type, nothing happens.

## Example

This code shows the card for the Linked data type in cell E5.

```vb
Range("E5").ShowCard
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]