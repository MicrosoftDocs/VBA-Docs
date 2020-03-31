---
title: InternetTimeout property example (VC++)
ROBOTS: INDEX
ms.prod: access
ms.assetid: 56c0d0df-8f8b-428f-ece9-ae5b98c9b820
ms.date: 06/08/2019
localization_priority: Normal
---


# InternetTimeout property example (VC++)

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [InternetTimeout](https://msdn.microsoft.com/library/66fc6e87-3d23-ce2c-18f5-0fc83ac43801%28Office.15%29.aspx) property, which exists on the [DataControl](https://msdn.microsoft.com/library/ac430669-7628-696c-c036-b5d35405d788%28Office.15%29.aspx) and [DataSpace](https://msdn.microsoft.com/library/7db181d5-422b-49fe-b6af-a20f5da520ff%28Office.15%29.aspx) objects. In this case, the **InternetTimeout** property is demonstrated on the **DataControl** object and the timeout is set to 20 seconds.

```cpp

// BeginInternetTimeoutCpp#import "c:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")#import "C:\Program Files\Common Files\System\MSADC\msadco.dll" 
#include <ole2.h>#include <stdio.h>
#include <conio.h> 
// Function declarationsinline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);};
void InternetTimeOutX(void);void PrintProviderError(_ConnectionPtr pConnection);
void PrintComError(_com_error &e); 
//////////////////////////////////////////////////////////// //
// Main Function //// //
////////////////////////////////////////////////////////// 
void main(){
if(FAILED(::CoInitialize(NULL)))return; 
InternetTimeOutX(); 
::CoUninitialize();} 
//////////////////////////////////////////////////////////// //
// InternetTimeOutX Function //// //
////////////////////////////////////////////////////////// 
void InternetTimeOutX(void){
HRESULT hr = S_OK; 
// Define ADO object pointers.// Initialize pointers on define.
// These are in the ADODB:: namespace._RecordsetPtr pRst = NULL; 
//Define RDS object pointersRDS::IBindMgrPtr dc ; 
try{
TESTHR(dc.CreateInstance(__uuidof(RDS::DataControl)));dc->Server = "https://MyServer";
dc->Connect = "Data Source='AuthorDatabase'";dc->SQL = "SELECT * FROM Authors"; 
// Wait at least 20 seconds.dc->InternetTimeout = 20000;
dc->Refresh(); 
// Use another Recordset as a conveniencepRst = dc->GetRecordset();
while(!(pRst->EndOfFile)){
printf("%s %s",(LPSTR) (_bstr_t) pRst->Fields->GetItem("au_fname")->Value,
(LPSTR) (_bstr_t) pRst->Fields->GetItem("au_lname")->Value); 
pRst->MoveNext();}
pRst->Close();} 
catch (_com_error &e){
PrintProviderError(pRst->GetActiveConnection());PrintComError(e);
}} 
//////////////////////////////////////////////////////////// //
// PrintProviderError Function //// //
////////////////////////////////////////////////////////// 
void PrintProviderError(_ConnectionPtr pConnection){
// Print Provider Errors from Connection object.// pErr is a record object in the Connection's Error collection.
ErrorPtr pErr = NULL; 
if( (pConnection->Errors->Count) > 0){
long nCount = pConnection->Errors->Count;// Collection ranges from 0 to nCount -1.
for(long i = 0; i < nCount; i++){
pErr = pConnection->Errors->GetItem(i);printf("\t Error number: %x\t%s", pErr->Number,
pErr->Description);}
}} 
//////////////////////////////////////////////////////////// //
// PrintComError Function //// //
////////////////////////////////////////////////////////// 
void PrintComError(_com_error &e){
_bstr_t bstrSource(e.Source());_bstr_t bstrDescription(e.Description()); 
// Print Com errors.printf("Error\n");
printf("\tCode = %08lx\n", e.Error());printf("\tCode meaning = %s\n", e.ErrorMessage());
printf("\tSource = %s\n", (LPCSTR) bstrSource);printf("\tDescription = %s\n", (LPCSTR) bstrDescription);
}// EndInternetTimeoutCpp
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]