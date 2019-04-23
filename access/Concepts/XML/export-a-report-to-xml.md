---
title: Export a report to XML
ms.prod: access
ms.assetid: 7e746a40-6227-1481-f631-702c3cf42d0f
ms.date: 09/26/2018
localization_priority: Normal
---


# Export a report to XML

This procedure does the following:

- Exports the Invoice report in the current database to an XML file. 
- Exports presentation information.
- Places images in the Images folder. 
- Exports the report to the default HTML wrapper. 
- Creates a file containing the ReportML list.


```vb
Private Sub ExportReport() 
 
 Const CREATE_REPORTML = 16 
 
 Application.ExportXML _ 
 ObjectType:=acExportReport, _ 
 DataSource:="Invoice", _ 
 DataTarget:="C:\Invoice.xml", _ 
 PresentationTarget:="C:\InvoiceReport.xsl", _ 
 ImageTarget:="C:\Images", _ 
 OtherFlags:=CREATE_REPORTML 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]