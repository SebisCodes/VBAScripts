VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ExportToSTEP(ByRef oDoc As Document, ByVal path As String)
    ' Get the STEP translator Add-In.
    Dim oSTEPTranslator As TranslatorAddIn
    Set oSTEPTranslator = ThisApplication.ApplicationAddIns.ItemById("{90AF7F40-0C01-11D5-8E83-0010B541CD80}")

    If oSTEPTranslator Is Nothing Then
        MsgBox "Could not access STEP translator."
        Exit Sub
    End If

    Dim oContext As TranslationContext
    Set oContext = ThisApplication.TransientObjects.CreateTranslationContext
    Dim oOptions As NameValueMap
    Set oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    If oSTEPTranslator.HasSaveCopyAsOptions(ThisApplication.ActiveDocument, oContext, oOptions) Then
        ' Set application protocol.
        ' 2 = AP 203 - Configuration Controlled Design
        ' 3 = AP 214 - Automotive Design
        oOptions.value("ApplicationProtocolType") = 3

        ' Other options...
        'oOptions.Value("Author") = ""
        'oOptions.Value("Authorization") = ""
        'oOptions.Value("Description") = ""
        'oOptions.Value("Organization") = ""

        oContext.Type = kFileBrowseIOMechanism

        Dim oData As DataMedium
        Set oData = ThisApplication.TransientObjects.CreateDataMedium
        oData.filename = path

        Call oSTEPTranslator.SaveCopyAs(oDoc, oContext, oOptions, oData)
    End If
End Sub
