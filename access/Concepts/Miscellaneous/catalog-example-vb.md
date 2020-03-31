---
title: Catalog example (VB)
ms.prod: access
ms.assetid: fcd982a8-0033-8643-99b0-a1b065e445f4
ms.date: 06/08/2019
localization_priority: Normal
---


# Catalog example (VB)

**Applies to:** Access 2013 | Access 2016

This Visual Basic project creates a new cube using MDX. Then, it documents the structure of a cube in a Microsoft Word document.


```vb
Sub cmdCreateDocForCube_Click() 
 On Error GoTo Error_cmdCreateDocForCube_Click 
 
 Dim cn As ADODB.Connection 
 Dim s As String 
 Dim strProvider As String 
 Dim strDataSource As String 
 Dim strSourceDSN As String 
 Dim strSourceDSNSuffix As String 
 Dim strCreateCube As String 
 Dim strInsertInto As String 
 
 '*----------------------------------------------------------------------- 
 '* The following code builds a cube file then documents the properties 
 '* with an OLE Connection to Word 8.0 
 '*----------------------------------------------------------------------- 
 
 '*----------------------------------------------------------------------- 
 '* Add Provider to the connection string. 
 '*----------------------------------------------------------------------- 
 strProvider = "PROVIDER=MSOLAP" 
 
 '*----------------------------------------------------------------------- 
 '* Add DataSource, the name of the file we will create. 
 '*----------------------------------------------------------------------- 
 
 strDataSource = "DATA SOURCE=c:\DocumentCube.cub" 
 
 '*----------------------------------------------------------------------- 
 '* Add Source DSN, the connection string for where the data comes from. 
 '* We need to quote the value so it is parsed as one value. 
 '* This can either be an ODBC connection string or an OLE DB connection 
 '* string. (As returned by the Data Source Locator component.) 
 '*----------------------------------------------------------------------- 
 strSourceDSN = "SOURCE_DSN=FoodMart 2000" 
 
 '*----------------------------------------------------------------------- 
 '* Add CREATE CUBE. This defines the structure of the cube, but not the 
 '* data in it. The BNF for is documented in the OLE DB for OLAP 
 '* Programmer's Reference. Note that we can quote names with either 
 '* double quotes or square brackets. 
 '*----------------------------------------------------------------------- 
 strCreateCube = "CREATECUBE=CREATE CUBE Sample( " 
 strCreateCube = strCreateCube & "DIMENSION [Product]," 
 strCreateCube = strCreateCube & "LEVEL [All Products] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Product Family] ," 
 strCreateCube = strCreateCube & "LEVEL [Product Department] ," 
 strCreateCube = strCreateCube & "LEVEL [Product Category] ," 
 strCreateCube = strCreateCube & "LEVEL [Product Subcategory] ," 
 strCreateCube = strCreateCube & "LEVEL [Brand Name] ," 
 strCreateCube = strCreateCube & "LEVEL [Product Name] ," 
 strCreateCube = strCreateCube & "DIMENSION [Store]," 
 strCreateCube = strCreateCube & "LEVEL [All Stores] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Store Country] ," 
 strCreateCube = strCreateCube & "LEVEL [Store State] ," 
 strCreateCube = strCreateCube & "LEVEL [Store City] ," 
 strCreateCube = strCreateCube & "LEVEL [Store Name] ," 
 strCreateCube = strCreateCube & "DIMENSION [Store Type]," 
 strCreateCube = strCreateCube & _ 
 "LEVEL [All Store Type] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Store Type] ," 
 strCreateCube = strCreateCube & "DIMENSION [Time] TYPE TIME," 
 strCreateCube = strCreateCube & "HIERARCHY [Column]," 
 strCreateCube = strCreateCube & "LEVEL [All Time] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Year] TYPE YEAR," 
 strCreateCube = strCreateCube & "LEVEL [Quarter] TYPE QUARTER," 
 strCreateCube = strCreateCube & "LEVEL [Month] TYPE MONTH," 
 strCreateCube = strCreateCube & "LEVEL [Week] TYPE WEEK," 
 strCreateCube = strCreateCube & "LEVEL [Day] TYPE DAY," 
 strCreateCube = strCreateCube & "HIERARCHY [Formula]," 
 strCreateCube = strCreateCube & _ 
 "LEVEL [All Formula Time] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Year] TYPE YEAR," 
 strCreateCube = strCreateCube & "LEVEL [Quarter] TYPE QUARTER," 
 strCreateCube = strCreateCube & _ 
 "LEVEL [Month] TYPE MONTH OPTIONS (SORTBYKEY) ," 
 strCreateCube = strCreateCube & "DIMENSION [Warehouse]," 
 strCreateCube = strCreateCube & _ 
 "LEVEL [All Warehouses] TYPE ALL," 
 strCreateCube = strCreateCube & "LEVEL [Country] ," 
 strCreateCube = strCreateCube & "LEVEL [State Province] ," 
 strCreateCube = strCreateCube & "LEVEL [City] ," 
 strCreateCube = strCreateCube & "LEVEL [Warehouse Name] ," 
 strCreateCube = strCreateCube & "MEASURE [Store Invoice] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#'," 
 strCreateCube = strCreateCube & "MEASURE [Supply Time] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#'," 
 strCreateCube = strCreateCube & "MEASURE [Warehouse Cost] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#'," 
 strCreateCube = strCreateCube & "MEASURE [Warehouse Sales] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#'," 
 strCreateCube = strCreateCube & "MEASURE [Units Shipped] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#'," 
 strCreateCube = strCreateCube & "MEASURE [Units Ordered] " 
 strCreateCube = strCreateCube & "Function Sum " 
 strCreateCube = strCreateCube & "Format '#.#')" 
 '*----------------------------------------------------------------------- 
 '* Add INSERT INTO. This defines where the data comes from, and how it 
 '* maps into the already-defined cube structure. Note that the SELECT 
 '* clause might just be passed through to the relational database. 
 '* So I could pass in a stored procedure, for example. If we needed to, 
 '* we could quote this whole thing. Note that the columns in the SELECT 
 '* can be in any order. One merely has to adjust the ordering of the 
 '* list of level/measure names to match the SELECT ordering. 
 '*----------------------------------------------------------------------- 
 strInsertInto = strInsertInto & _ 
 "INSERTINTO=INSERT INTO Sample( " & _ 
 "Product.[Product Family], Product.[Product Department]," 
 strInsertInto = strInsertInto & _ 
 "Product.[Product Category], Product.[Product Subcategory]," 
 strInsertInto = strInsertInto & _ 
 "Product.[Brand Name], Product.[Product Name]," 
 strInsertInto = strInsertInto & _ 
 "Store.[Store Country], Store.[Store State], Store.[Store City]," 
 strInsertInto = strInsertInto & _ 
 "Store.[Store Name], [Store Type].[Store Type], [Time].[Column]," 
 strInsertInto = strInsertInto & _ 
 "[Time].Formula.Year, [Time].Formula.Quarter, " & _ 
 "[Time].Formula.Month.[Key]," 
 strInsertInto = strInsertInto & _ 
 "[Time].Formula.Month.Name, Warehouse.Country, " & _ 
 "Warehouse.[State Province]," 
 strInsertInto = strInsertInto & _ 
 "Warehouse.City, Warehouse.[Warehouse Name], " & _ 
 "Measures.[Store Invoice]," 
 strInsertInto = strInsertInto & _ 
 "Measures.[Supply Time], Measures.[Warehouse Cost], " & _ 
 "Measures.[Warehouse Sales]," 
 strInsertInto = strInsertInto & _ 
 "Measures.[Units Shipped], Measures.[Units Ordered] )" 
 '*----------------------------------------------------------------------- 
 '* Add some options to the INSERT INTO if we need to. These can control 
 '* if the SELECT clause is analyzed or just passed through, and if the 
 '* storage mode is MOLAP or ROLAP (DEFER_DATA). 
 '* strInsertInto = strInsertInto & " OPTIONS ATTEMPT_ANALYSIS" 
 '*----------------------------------------------------------------------- 
 
 '*----------------------------------------------------------------------- 
 '* Add the SELECT clause of the INSERT INTO statement. Note that it is 
 '* merely concatenated onto the end of the INSERT INTO statement. OLAP 
 '* Services will pass this through to the source database if unable to 
 '* parse it. Note that for OLAP Services to analyze the SELECT clause, 
 '* each column must be qualified with the table name. 
 '*----------------------------------------------------------------------- 
 
 strInsertInto = strInsertInto & _ 
 "SELECT product_class.product_family AS Col1," 
 strInsertInto = strInsertInto & _ 
 "product_class.product_department AS Col2," 
 strInsertInto = strInsertInto & "product_class.product_category AS Col3," 
 strInsertInto = strInsertInto & _ 
 "product_class.product_subcategory AS Col4," 
 strInsertInto = strInsertInto & "product.brand_name AS Col5," 
 strInsertInto = strInsertInto & "product.product_name AS Col6," 
 strInsertInto = strInsertInto & "store.store_country AS Col7," 
 strInsertInto = strInsertInto & "store.store_state AS Col8," 
 strInsertInto = strInsertInto & "store.store_city AS Col9," 
 strInsertInto = strInsertInto & "store.store_name AS Col10," 
 strInsertInto = strInsertInto & "store.store_type AS Col11," 
 strInsertInto = strInsertInto & "time_by_day.the_date AS Col12," 
 strInsertInto = strInsertInto & "time_by_day.the_year AS Col13," 
 strInsertInto = strInsertInto & "time_by_day.quarter AS Col14," 
 strInsertInto = strInsertInto & "time_by_day.month_of_year AS Col15," 
 strInsertInto = strInsertInto & "time_by_day.the_month AS Col16," 
 strInsertInto = strInsertInto & "warehouse.warehouse_country AS Col17," 
 strInsertInto = strInsertInto & _ 
 "warehouse.warehouse_state_province AS Col18," 
 strInsertInto = strInsertInto & "warehouse.warehouse_city AS Col19," 
 strInsertInto = strInsertInto & "warehouse.warehouse_name AS Col20," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.store_invoice AS Col21," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.supply_time AS Col22," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.warehouse_cost AS Col23," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.warehouse_sales AS Col24," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.units_shipped AS Col25," 
 strInsertInto = strInsertInto & _ 
 "inventory_fact_1997.units_ordered AS Col26 " 
 strInsertInto = strInsertInto & _ 
 "From [inventory_fact_1997], [product], [product_class], " & _ 
 "[time_by_day], [store], [warehouse] " 
 strInsertInto = strInsertInto & _ 
 "Where [inventory_fact_1997].[product_id] = [product]." & _ 
 "[product_id] And " 
 strInsertInto = strInsertInto & _ 
 "[product].[product_class_id] = [product_class]." & _ 
 "[product_class_id] And " 
 strInsertInto = strInsertInto & _ 
 "[inventory_fact_1997].[time_id] = [time_by_day].[time_id] And " 
 strInsertInto = strInsertInto & _ 
 "[inventory_fact_1997].[store_id] = [store].[store_id] And " 
 strInsertInto = strInsertInto & _ 
 "[inventory_fact_1997].[warehouse_id] = [warehouse].[warehouse_id]" 
 
 '*----------------------------------------------------------------------- 
 '* Create the cube by passing connection string to Open. 
 '*----------------------------------------------------------------------- 
 
 Set cn = New ADODB.Connection 
 s = strProvider & ";" & strDataSource & ";" & strSourceDSN & ";" & _ 
 strCreateCube & ";" & strInsertInto & ";" 
 Screen.MousePointer = vbHourglass 
 cn.Open s 
 
 '*----------------------------------------------------------------------- 
 '* Cube file is written to hard drive a Word Document can be produced by 
 '* automating Word with VB 
 '*----------------------------------------------------------------------- 
 
 Dim cat As New ADOMD.Catalog 
 Dim cdf As ADOMD.CubeDef 
 Dim i As Integer 
 Dim di As Integer 
 Dim hi As Integer 
 Dim le As Integer 
 Dim mem As Integer 
 Dim docWord As Word.Document 
 Dim rngCurrent As Word.Range 
 Dim SenCount As Integer 
 Dim strServer As String 
 Dim strSource As String 
 Dim strCubeName As String 
 Dim appWord As Object 
 '*----------------------------------------------------------------------- 
 '* Connection is made to cube file 
 '*----------------------------------------------------------------------- 
 cat.ActiveConnection = "DATA SOURCE=c:\DocumentCube.cub;Provider=msolap;" 
 
 '*----------------------------------------------------------------------- 
 '* Cube Definition is set to Name of Cube in cube file 
 '*----------------------------------------------------------------------- 
 Set cdf = cat.CubeDefs("Sample") 
 
 '*----------------------------------------------------------------------- 
 '* Object is created to hold Word 
 '*----------------------------------------------------------------------- 
 Set appWord = CreateObject("Word.Application") 
 
 '*----------------------------------------------------------------------- 
 '* Create the document variable 
 '*----------------------------------------------------------------------- 
 Set docWord = appWord.Documents.Add() 
 
 Set rngCurrent = docWord.Content 
 
 SenCount = 0 
 
 '*----------------------------------------------------------------------- 
 '* Cube Title and Header written to Document 
 '*----------------------------------------------------------------------- 
 With rngCurrent 
 .InsertAfter "Report for Sample Cube" 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = True 
 docWord.Paragraphs(SenCount).Range.Underline = wdUnderlineSingle 
 docWord.Paragraphs(SenCount).Range.Italic = False 
 docWord.Paragraphs(SenCount).Range.Font.Size = 18 
 
 '*----------------------------------------------------------------------- 
 '* Properties of Cube are written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Properties of Cube are written to Document" 
 For i = 0 To cdf.Properties.Count - 1 
 .InsertAfter "(" & i & ") " & cdf.Properties(i).Name & ": " & _ 
 cdf.Properties(i).Value 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = False 
 docWord.Paragraphs(SenCount).Range.Italic = True 
 docWord.Paragraphs(SenCount).Range.Font.Size = 8 
 Next i 
 
 '*----------------------------------------------------------------------- 
 '* Dimension Name(s) written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Dimension Name(s) written to Document" 
 For di = 0 To cdf.Dimensions.Count - 1 
 .InsertAfter "Dimension: " & cdf.Dimensions(di).Name 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = True 
 docWord.Paragraphs(SenCount).Range.Italic = False 
 docWord.Paragraphs(SenCount).Range.Font.Size = 14 
 
 '*----------------------------------------------------------------------- 
 '* Properties of Dimension are written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Properties of Dimension are written to Document" 
 For i = 0 To cdf.Dimensions(di).Properties.Count - 1 
 .InsertAfter "(" & i & ") " & _ 
 cdf.Dimensions(di).Properties(i).Name & ": " & _ 
 cdf.Dimensions(di).Properties(i).Value 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = False 
 docWord.Paragraphs(SenCount).Range.Italic = True 
 docWord.Paragraphs(SenCount).Range.Font.Size = 8 
 Next i 
 
 '*----------------------------------------------------------------------- 
 '* Hierarchy Name(s) written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Hierarchy Name(s) written to Document" 
 For hi = 0 To cdf.Dimensions(di).Hierarchies.Count - 1 
 .InsertAfter vbTab & "Hierarchy: " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Name 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = True 
 docWord.Paragraphs(SenCount).Range.Italic = False 
 docWord.Paragraphs(SenCount).Range.Font.Size = 12 
 
 '*----------------------------------------------------------------------- 
 '* Properties of Hierarchy are written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Properties of Hierarchy are written to Document" 
 For i = 0 To cdf.Dimensions(di).Hierarchies(hi).Properties.Count - 1 
 .InsertAfter vbTab & "(" & i & ") " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Properties(i).Name & _ 
 ": " & cdf.Dimensions(di).Hierarchies(hi).Properties(i).Value 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = False 
 docWord.Paragraphs(SenCount).Range.Italic = True 
 docWord.Paragraphs(SenCount).Range.Font.Size = 8 
 Next i 
 
 '*----------------------------------------------------------------------- 
 '* Level Name(s) written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Level Name(s) written to Document" 
 For le = 0 To cdf.Dimensions(di).Hierarchies(hi).Levels.Count - 1 
 .InsertAfter vbTab & vbTab & "Level: " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Levels(le).Name & _ 
 " with a Member Count of: " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Levels(le).Members.Count 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = True 
 docWord.Paragraphs(SenCount).Range.Italic = False 
 docWord.Paragraphs(SenCount).Range.Font.Size = 10 
 
 '*----------------------------------------------------------------------- 
 '* Properties of Level are written to Document 
 '*----------------------------------------------------------------------- 
 Debug.Print "Properties of Level are written to Document" 
 For i = 0 To _ 
 cdf.Dimensions(di).Hierarchies(hi).Levels(le).Properties.Count - 1 
 .InsertAfter vbTab & vbTab & "(" & i & ") " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Levels(le).Properties(i).Name & ": " & _ 
 cdf.Dimensions(di).Hierarchies(hi).Levels(le).Properties(i).Value 
 .InsertAfter vbCrLf 
 SenCount = SenCount + 1 
 docWord.Paragraphs(SenCount).Range.Bold = False 
 docWord.Paragraphs(SenCount).Range.Italic = True 
 docWord.Paragraphs(SenCount).Range.Font.Size = 8 
 Next i 
 
 Next le 
 Next hi 
 Next di 
 
 '*----------------------------------------------------------------------- 
 '* Set Word Document to visible 
 '*----------------------------------------------------------------------- 
 Debug.Print "Set Word Document to visible" 
 appWord.Visible = True 
 End With 
 
 Screen.MousePointer = vbDefault 
 
 '*----------------------------------------------------------------------- 
 '* Set Word Object to nothing to drop OLE connection 
 '*----------------------------------------------------------------------- 
 Set appWord = Nothing 
Exit Sub 
 
Error_cmdCreateDocForCube_Click: 
 cn.Cancel 
 Set cn = Nothing 
 Screen.MousePointer = vbDefault 
 MsgBox Err.Description 
End Sub 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]