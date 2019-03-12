---
title: DoCmd.SendObject method (Access)
keywords: vbaac10.chm4180
f1_keywords:
- vbaac10.chm4180
ms.prod: access
api_name:
- Access.DoCmd.SendObject
ms.assetid: 881004c6-2dd7-55f1-2a16-2d28034125a8
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.SendObject method (Access)

The **SendObject** method carries out the SendObject action in Visual Basic.


## Syntax

_expression_.**SendObject** (_ObjectType_, _ObjectName_, _OutputFormat_, _To_, _Cc_, _Bcc_, _Subject_, _MessageText_, _EditMessage_, _TemplateFile_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**[AcSendObjectType](Access.AcSendObjectType.md)**|An **AcSendObjectType** constant that specifies the type of object to send.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _ObjectType_ argument. If you want to include the active object in the mail message, specify the object's type with the _ObjectType_ argument and leave this argument blank. If you leave both the _ObjectType_ and _ObjectName_ arguments blank (the default constant, **acSendNoObject**, is assumed for the _ObjectType_ argument), Microsoft Access sends a message to the electronic mail application without an included database object.<br/><br/>If you run Visual Basic code containing the **SendObject** method in a library database, Access looks for the object with this name first in the library database, and then in the current database.|
| _OutputFormat_|Optional|**Variant**|A constant that specifies the format in which to send the object. Possible values include **acFormatHTML**, **acFormatRTF**, **acFormatSNP**, **acFormatTXT**, **acFormatXLS**, **acFormatXLSB**, **acFormatXLSX**, **acFormatXPS**, and **acFormatPDF**.|
| _To_|Optional|**Variant**|A string expression that lists the recipients whose names you want to put on the To line in the mail message. Separate the recipient names that you specify in this argument and in the _Cc_ and _Bcc_ arguments with a semicolon (;), or with the list separator set on the **Number** tab of the **Regional Settings Properties** dialog box in the Windows Control Panel. If the recipient names aren't recognized by the mail application, the message isn't sent and an error occurs. If you leave this argument blank, Microsoft Access prompts you for the recipients.|
| _Cc_|Optional|**Variant**|A string expression that lists the recipients whose names you want to put on the **Cc** line in the mail message. If you leave this argument blank, the **Cc** line in the mail message is blank.|
| _Bcc_|Optional|**Variant**|A string expression that lists the recipients whose names you want to put on the **Bcc** line in the mail message. If you leave this argument blank, the **Bcc** line in the mail message is blank.|
| _Subject_|Optional|**Variant**|A string expression containing the text that you want to put on the **Subject** line in the mail message. If you leave this argument blank, the **Subject** line in the mail message is blank.|
| _MessageText_|Optional|**Variant**|A string expression containing the text that you want to include in the body of the mail message, after the object. If you leave this argument blank, the object is all that's included in the body of the mail message.|
| _EditMessage_|Optional|**Variant**|Use **True** (1) to open the electronic mail application immediately with the message loaded, so the message can be edited. Use **False** (0) to send the message without editing it. If you leave this argument blank, the default (**True**) is assumed.|
| _TemplateFile_|Optional|**Variant**|A string expression that's the full name, including the path, of the file that you want to use as a template for an HTML file.|

## Remarks

You can use the SendObject action to include the specified Microsoft Access datasheet, form, report, or module in an electronic mail message, where it can be viewed and forwarded. You can include objects in Microsoft Excel 2000 (\*.xls), MS-DOS text (\*.txt), rich-text (\*.rtf), or HTML (\*.html) format in messages for Microsoft Outlook, Microsoft Exchange, or another electronic mail application that uses the Mail Applications Programming Interface (MAPI).

The following rules apply when you use the SendObject action to include a database object in a mail message:

- You can send table, query, and form datasheets. In the included object, all fields in the datasheet look as they do in Access, except fields containing OLE objects. The columns for these fields are included in the object, but the fields are blank.
    
- For a control bound to a **Yes/No** field (a toggle button, option button, or check box), the output file displays the value 1 (Yes) or 0 (No).
    
- For a text box bound to a **Hyperlink** field, the output file displays the hyperlink for all output formats except MS-DOS text (in this case, the hyperlink is just displayed as normal text).
    
- If you send a form in Form view, the included object always contains the form's Datasheet view.
    
- If you send a report, the only controls that are included in the object are text boxes (for .xls files), or text boxes and labels (for .rtf, .txt, and .html files). All other controls are ignored. Header and footer information is also not included. The only exception to this is that when you send a report in Excel format, a text box in a group footer containing an expression with the **Sum** function is included in the object. No other control in a header or footer (and no aggregate function other than **Sum**) is included in the object.
    
- Subreports are included in the object. Subforms are included when outputting to .asp, but only when outputting as a form (not a datasheet).
    
- When you send a datasheet, form, or data access page in HTML format, one .html file is created. When you send a report in HTML format, one .html file is created for each page in the report.
    
Modules can be sent only in MS-DOS Text format, so if you specify **acSendModule** for the _ObjectType_ argument, you must specify **acFormatTXT** for the _OutputFormat_ argument.

> [!NOTE]
> You can save as a PDF or XPS file from a 2007 Microsoft Office system program only after you install an add-in. For more information, see [2007 Microsoft Office Add-in: Microsoft Save as PDF or XPS](https://www.microsoft.com/download/details.aspx?id=7).
 

## Example

The following code example includes the **Employees** table in a mail message in Microsoft Excel format and specifies **To**, **Cc**, and **Subject** lines in the mail message. The mail message is sent immediately, without editing.

```vb
DoCmd.SendObject acSendTable, "Employees", acFormatXLS, _ 
    "Nancy Davolio; Andrew Fuller", "Joan Weber", , _ 
    "Current Spreadsheet of Employees", , False
```

<br/>

The following example shows how to create an email message with Microsoft Outlook and display it to the user.

```vb
Public Function CreateEmailWithOutlook( _
    MessageTo As String, _
    Subject As String, _
    MessageBody As String)

    ' Define app variable and get Outlook using the "New" keyword
    Dim olApp As New Outlook.Application
    Dim olMailItem As Outlook.MailItem  ' An Outlook Mail item
 
    ' Create a new email object
    Set olMailItem = olApp.CreateItem(olMailItem)

    ' Add the To/Subject/Body to the message and display the message
    With olMailItem
        .To = MessageTo
        .Subject = Subject
        .Body = MessageBody
        .Display    ' To show the email message to the user
    End With

    ' Release all object variables
    Set olMailItem = Nothing
    Set olApp = Nothing

End Function
```

<br/>

The following example shows how to create an email message with Microsoft Outlook and send it without displaying the email message to the user.

```vb
Public Function SendEmailWithOutlook( _
    MessageTo As String, _
    Subject As String, _
    MessageBody As String)

    ' Define app variable and get Outlook using the "New" keyword
    Dim olApp As New Outlook.Application
    Dim olMailItem As Outlook.MailItem  ' An Outlook Mail item
 
    ' Create a new email object
    Set olMailItem = olApp.CreateItem(olMailItem)

    ' Add the To/Subject/Body to the message and display the message
    With olMailItem
        .To = MessageTo
        .Subject = Subject
        .Body = MessageBody
        .Send       ' Send the message immediately
    End With

    ' Release all object variables
    Set olMailItem = Nothing
    Set olApp = Nothing

End Function
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
