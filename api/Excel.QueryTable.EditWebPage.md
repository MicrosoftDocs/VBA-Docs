---
title: QueryTable.EditWebPage property (Excel)
keywords: vbaxl10.chm518130
f1_keywords:
- vbaxl10.chm518130
ms.prod: excel
api_name:
- Excel.QueryTable.EditWebPage
ms.assetid: 4de607d1-266f-cbd4-c236-af748cfe0d03
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.EditWebPage property (Excel)

Returns or sets the webpage Uniform Resource Locator (URL) for a web query. Read/write **Variant**.


## Syntax

_expression_.**EditWebPage**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

The **EditWebPage** property returns **null** if not set. The **EditWebPage** property is only meaningful if the query type is Web or OLE.

If the **EditWebPage** is not **null**, ignore the **[WebTables](Excel.QueryTable.WebTables.md)** property for refreshing. As a result, an XML query and the **WebTables** property refers to the table in the original webpage and should only be used in the edit case to pre-populate the **Web Query** dialog box.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **EditWebPage** property applies only to **QueryTable** objects.


## Example

In this example, Microsoft Excel displays to the user a webpage URL. This example assumes a **QueryTable** object in cell A1 exists on the active worksheet and that a file called MyHomepage.htm exists on the C:\ drive.

```vb
Sub ReturnURL() 
 
 ' Set the EditWebPage property to a source. 
 Range("A1").QueryTable.EditWebPage = "C:\MyHomepage.htm" 
 
 ' Display the source to the user. 
 MsgBox Range("A1").QueryTable.EditWebPage 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]