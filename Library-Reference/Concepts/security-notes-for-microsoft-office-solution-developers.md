---
title: Security notes for Office solution developers
ms.prod: office
ms.assetid: 076ce284-5d1d-4823-ba74-f5e5c05bae9b
ms.date: 01/02/2019
localization_priority: Normal
---


# Security notes for Office solution developers

## Setting Office security in a testing environment

> [!NOTE] 
> You can include Visual Basic for Applications (VBA) code or run COM add-ins only in a macro-enabled document, worksheet, or presentation. You can create a macro-enabled file by saving the documents with a .docm or .dotm extension in Word; an .xlsm, .xltm, or .xlam extension in Excel; or a .pptm, .potm, .ppam, or .ppsm extension in PowerPoint.

To install and run an unsigned COM add-in, the **Require Application Add-ins to be signed by Trusted Publisher** and the **Disable all Application Add-ins** options must be cleared on the **Add-ins** tab in the Trust Center. To open the **Add-ins** tab, choose the **File** tab, and then choose **Options** > **Trust Center** > **Trust Center Settings**> **Add-ins**. 

To run all VBA macros, including those that have not been digitally signed, the **Enable all macros** option must be selected in the Trust Center. To view the **Macro Settings** options, choose the **File** tab, and then choose **Options** > **Trust Center** > **Trust Center Settings** > **Macros Settings**. For security reasons, it is strongly recommended that you do this only in a testing environment. After you complete your testing, set the options back to their original state.

On the **Macro Settings** tab of the Trust Center, you can also set options to **Disable all macros without notification**, **Disable all macros with notification**, or **Disable all macro except digitally signed macros**. You can also disable macros by saving the Word document, Excel worksheet, or PowerPoint presentation as macro-disabled files (.docm, .xlsm, or .pptm, respectively). You can also set or disable access to the VBA project object model from the **Macro Settings** tab by selecting or clearing the **Trust access to the VBA project object model** option.

> [!NOTE] 
> On the Office Fluent user interface ribbon, when COM and application-specific add-ins are enabled and loaded, their controls are displayed on an **Add-ins** tab.

You can see a list of available add-ins on the **Add-ins** tab in the Trust Center. On the same tab, you can enable, disable, add, or remove COM or Word add-ins by selecting the type of add-in in the drop-down box next to the **Manage** label, and then choosing the **Go** button.


## Modifying the Windows registry

Modifying the Windows registry in any manner, whether in the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is always a good practice to back up a computer's registry first before you modify it. If you are running Microsoft Windows NT, Windows 2000, Windows XP, or Windows Server 2003, you should also update your Emergency Repair Disk (ERD).

For information about how to edit the registry, see the "Changing Keys and Values" Help topic in the Registry Editor (Regedit.exe) or the "Add and Delete Information in the Registry" and "Edit Registry Information" topics in the Registry Editor (Regedt32.exe).


## Making Windows API function calls

Before calling Windows Application Programming Interface (API) functions, you should understand how arguments and data types are handled by the Windows API DLLs. Incorrectly calling Windows functions might result in invalid page faults or other unexpected behaviors. For more information about how to call Windows functions, see the topic "The Windows API and Other Dynamic-Link Libraries" in the Office 2000 Developer Online Documentation.


## Digitally signing code

Digitally signing a document is the process of "stamping" a document so that the recipient of the document can be assured that it came from a particular source, and can detect whether the contents of the document have changed since the document was signed. Additionally, digital signatures can be used to mark a document as read-only to protect its authenticity and integrity.

In addition to digital signatures, documents can also contain in-document signatures that are visible in the document's content. The originator of the document can create unsigned documents with signature lines that can be transmitted to the recipient to sign. The recipient opens the document, locates the signature line, signs the document, and then returns it to the sender.

Basically, the steps to digitally sign a document include: 

1. The document's originator compacts the document's content into a few lines by using a process called "hashing." The compressed content is called a message digest. Hashing is performed by software that is created for that purpose.
    
2. The document's originator then encrypts the message digest by using a private key obtained from a signing authority. The result is a digital signature. 
    
3. The originator attaches the digital signature to the document. All of the data that was hashed has now been signed, and the signature has been encrypted and attached to the document.
    
4. The originator then sends the document to the recipient.
    
5. The recipient first decrypts the document by using a public key received from the originator. This changes the signature back to a message digest. If this works, it proves that the document was signed by the originator.
    
6. The recipient, using digital signing software, hashes the document into a message digest and compares this hash to the hash from the sender. If they match, this verifies that the contents of the document have not changed since the document was sent by the originator.
    
Digital signatures have been available to customers since Office XP. However, Office 2007 added features that make it easier for users to digitally sign documents, sign their documents to make them read-only, and add inline-document signature lines to a document. Office users can perform these tasks from the Office user interface that is available on the **File** tab.

Office 2007 also introduced members that make it easier to work with in-line signatures and digital signatures programmatically. 


## Deploying managed COM add-ins in Office securely

To comply with Office security, managed COM add-ins (COM add-ins targeting the common language runtime) must be digitally signed, and users' security settings should be set in the Office Trust Center to allow add-ins in your Office applications. Additionally, you must incorporate into your managed COM add-in project a small unmanaged proxy called a _shim_ to avoid unexpected security warnings. 


## Automating the Visual Basic Editor

In Office, when calling the features of the Visual Basic for Applications Extensibility object model, you might receive an error message that programmatic access to the Visual Basic project is not trusted. To prevent this message from appearing, choose the **File** tab > **Options** > **Trust Center** tab, and then choose **Trust Center Settings**. Next, on the **Macro Settings** tab, choose the **Trust access to the VBA project object model** box. By checking this box, you make it possible for macros in any macro-enabled documents that you open to access the core Visual Basic objects, methods, and properties. 

Setting the option represents a possible security hazard. The recommended behavior is to check the **Trust access to the VBA project object model** box only for the duration of a macro that accesses the Visual Basic object model. Make sure that you clear the **Trust access to the VBA project object model** box after the macro has finished running.


## Using passwords

Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code.

Always use strong passwords. Strong passwords should contain:

- Both lowercase and uppercase characters
- Numbers 
- Symbols (such as #, $, %, and ^) 
- At least eight characters
    
Strong passwords should not contain patterns, themes, or words found in a dictionary.

Examples of strong passwords include:

- `$tR0n9p@$s`   
- `G80dn[s$M4!`
    
> [!NOTE] 
> You should change your password frequently; for example, every one to three months.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]