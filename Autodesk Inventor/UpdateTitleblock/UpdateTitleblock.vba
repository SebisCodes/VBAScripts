Option Explicit

' A VBA-script tested on Inventor 2022 to update the titleblock in a drawing
'
' Contact: sebiscodes@gmail.com
' My Github: https://github.com/SebisCodes

Sub updateTitleblockPWT()
    Dim oDoc As DrawingDocument
    Set oDoc = ThisApplication.ActiveDocument

    If (oDoc.DocumentType <> kDrawingDocumentObject) Then Exit Sub
    On Error Resume Next
    
    Dim oTemplate As DrawingDocument
    Dim oSourceTitleBlockDef As TitleBlockDefinition
    Dim oSourceBorderDef As BorderDefinition
    Dim oNewTitleBlockDef As TitleBlockDefinition
    Dim oNewBorderDef As BorderDefinition
    Dim oSheet As sheet
    Dim oTitleBlockDefinition As TitleBlockDefinition
    Dim oBorderDefinition As BorderDefinition
    
    ' Get template titleblock
    ' !!! Replace 'ThisApplication.FileOptions.TemplatesPath & "Norm.idw"' with your template drawing file !!!
    Set oTemplate = ThisApplication.Documents.Open(ThisApplication.FileOptions.TemplatesPath & "Norm.idw", False) 
    Set oSourceTitleBlockDef = oTemplate.ActiveSheet.TitleBlock.Definition
    Set oSourceBorderDef = oTemplate.ActiveSheet.Border.Definition

    ' Iterate through the sheets.
    For Each oSheet In oDoc.Sheets
        oSheet.Activate
        If Not oSheet.TitleBlock Is Nothing Then
            oSheet.TitleBlock.Delete
        End If
        If Not oSheet.Border Is Nothing Then
            oSheet.Border.Delete
        End If
    Next
    
    For Each oTitleBlockDefinition In oDoc.TitleBlockDefinitions
        oTitleBlockDefinition.Delete
    Next
    
    For Each oBorderDefinition In oDoc.BorderDefinitions
        oBorderDefinition.Delete
    Next
    
    Set oNewBorderDef = oSourceBorderDef.CopyTo(oDoc)
    Set oNewTitleBlockDef = oSourceTitleBlockDef.CopyTo(oDoc)
    For Each oSheet In oDoc.Sheets
        oSheet.Activate
        Call oSheet.AddTitleBlock(oNewTitleBlockDef)
        If oNewBorderDef Is Nothing Then
            Call oSheet.AddDefaultBorder
        Else
            Call oSheet.AddBorder(oNewBorderDef)
        End If
    Next

    oTemplate.Close
End Sub
