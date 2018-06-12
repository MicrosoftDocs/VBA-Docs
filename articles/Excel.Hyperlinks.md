---
title: Hyperlinks Object (Excel)
keywords: vbaxl10.chm533072
f1_keywords:
- vbaxl10.chm533072
ms.prod: excel
api_name:
- Excel.Hyperlinks
ms.assetid: de28e0af-7a4c-56c3-5fe5-ac47d1654628
ms.date: 06/08/2017
---


# Hyperlinks Object (Excel)

Represents the collection of hyperlinks for a worksheet or range.


## Remarks

 Each hyperlink is represented by a **[Hyperlink](Excel.Hyperlink.md)** object.


## Example

Use the  **[Hyperlinks](Excel.Worksheet.Hyperlinks.md)** property to return the **Hyperlinks** collection. The following example checks the hyperlinks on worksheet one for a link that contains the word Microsoft.


```
For Each h in Worksheets(1).Hyperlinks 
 If Instr(h.Name, "Microsoft") <> 0 Then h.Follow 
Next
```

Use the  **[Add](Excel.Hyperlinks.Add.md)** method to create a hyperlink and add it to the **Hyperlinks** collection. The following example creates a new hyperlink for cell E5.




```
With Worksheets(1) 
 .Hyperlinks.Add .Range("E5"), "http://example.microsoft.com" 
End With
```


## Methods



|**Name**|
|:-----|
|[Add](Excel.Hyperlinks.Add.md)|
|[Delete](Excel.Hyperlinks.Delete.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Hyperlinks.Application.md)|
|[Count](Excel.Hyperlinks.Count.md)|
|[Creator](Excel.Hyperlinks.Creator.md)|
|[Item](Excel.Hyperlinks.Item.md)|
|[Parent](hyperlinks-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
