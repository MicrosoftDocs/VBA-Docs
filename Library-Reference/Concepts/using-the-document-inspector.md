---
title: Using the Document Inspector
ms.prod: office
ms.assetid: 62180311-ee41-1812-797d-3b5814add284
ms.date: 01/02/2019
localization_priority: Normal
---


# Using the Document Inspector

The **Document Inspector** gives users an easy way to examine documents for personal or sensitive information, text phrases, and other document contents. They can use the **Document Inspector** to remove unwanted information; for example, before distributing a document.

> [!NOTE] 
> Microsoft does not support the automatic removal of hidden information for signed or protected documents, or for documents that use Information Rights Management (IRM). We recommend that you run the **Document Inspector** before you sign a document or invoke IRM on a document.

As a developer, you can use the Document Inspector framework to extend the built-in modules and integrate your extensions into the standard user interface. 

The **Document Inspector** in Word, Excel, and PowerPoint includes the following enhancements.

## Built-in Document Inspector modules

The **Document Inspector** has modules that help users inspect and fix specific elements of a given document. The **Document Inspector** includes the following built-in modules.

### For all Office documents

- Embedded documents   
- OLE objects and packages 
- Data models 
- Content apps 
- Task Pane apps 
- Macros and VBA modules
- Legacy macros (XLM and WordBasic)
    
### For Excel documents

- PivotTables and slicers 
- PivotCharts
- Cube formulas
- Timelines (cache)
- Custom XML data
- Comments and annotations
- Document properties and personal information
- Headers and footers
- Hidden rows and columns   
- Hidden worksheets and names   
- Invisible content   
- External links and data functions   
- Excel surveys   
- Custom worksheet properties
    
### For PowerPoint documents

- Comments and annotations   
- Document properties and personal information   
- Invisible on-slide content   
- Off-slide content 
- Presentation notes
    
### For Word documents

- Comments, revisions, versions, and annotations 
- Document properties and personal information; this includes metadata, SharePoint properties, custom properties, and other content information  
- Custom XML data   
- Headers, footers, and watermarks   
- Invisible content  
- Hidden text
    

## Opening the Document Inspector

To open the **Document Inspector**:

1. Choose the **File** tab, and then choose **Info**.
    
2. Choose **Check for Issues**.
    
3. Choose **Inspect Document**.
    
Use the **Document Inspector** dialog box to select the type or types of data to find in the document.

After the modules complete the inspection, the **Document Inspector** displays the results for each module in a dialog box. If a given module finds data, the dialog box includes a **Remove All** button that you can click to remove that data. If the module does not find data, the dialog box displays a message to that effect.

If you choose to remove the data for a given module, the dialog box displays descriptive text that indicates whether the operation was successful or not. If the **Document Inspector** encounters errors during the operation, the module is flagged, displays an error message, and the data for that module does not change.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
