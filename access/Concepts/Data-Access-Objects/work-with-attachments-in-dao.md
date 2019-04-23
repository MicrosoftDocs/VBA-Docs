---
title: Work with attachments in DAO
ms.prod: access
ms.assetid: e175a47a-4d97-b93b-c152-809314ac5ba0
ms.date: 09/21/2018
localization_priority: Normal
---


# Work with attachments in DAO

In DAO, Attachment fields function just like other multi-valued fields. The field that contains the attachment contains a recordset that is a child to the table's recordset. There are two new DAO methods, **[LoadFromFile](../../../api/overview/Access.md)** and **[SaveToFile](../../../api/overview/Access.md)**, that deal exclusively with attachments.


## Add an attachment to a record

The **LoadFromFile** method loads a file from disk and adds the file as an attachment to the specified record. The following code example shows the syntax of the **LoadFromFile** method.


```vb
Recordset.Fields("FileData").LoadFromFile(<filename>)
```

> [!NOTE] 
> The **FileData** field is reserved internally by the Access database engine to store the binary attachment data.

The following code example uses the **LoadFromFile** method to load an employee's picture from disk.


```vb
   '  Instantiate the parent recordset.  
   Set rsEmployees = db.OpenRecordset("Employees") 
  
   … Code to move to desired employee 
  
   ' Activate edit mode. 
   rsEmployees.Edit 
  
   ' Instantiate the child recordset. 
   Set rsPictures = rsEmployees.Fields("Pictures").Value  
  
   ' Add a new attachment. 
   rsPictures.AddNew 
   rsPictures.Fields("FileData").LoadFromFile "EmpPhoto39392.jpg" 
   rsPictures.Update 
  
   ' Update the parent record 
   rsEmployees.Update 

```


## Save an attachment to disk

The following code example shows how to use the **SaveToFile** method to save all of the attachments for a specific employee to disk.


```vb
'  Instantiate the parent recordset.  
   Set rsEmployees = db.OpenRecordset("Employees") 
  
   … Code to move to desired employee 
  
   ' Instantiate the child recordset. 
   Set rsPictures = rsEmployees.Fields("Pictures").Value  
 
   '  Loop through the attachments. 
   While Not rsPictures.EOF 
  
      '  Save current attachment to disk in the "My Documents" folder. 
      rsPictures.Fields("FileData").SaveToFile _ 
                  "C:\Documents and Settings\Username\My Documents" 
      rsPictures.MoveNext 
   Wend 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
