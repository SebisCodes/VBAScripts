Option Explicit

' A VBA-script tested on Inventor 2022 to convert an assembly to a part with no link to its assembly
'
' Contact: sebiscodes@gmail.com
' My Github: https://github.com/SebisCodes

Sub assemblyToPart()

    Dim run As Integer
    run = MsgBox("Convert assembly to part?", vbOKCancel)
    If Not run = 1 Then
        End
    End If
    
    If (ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject) Then
        MsgBox ("Error! Active document must be an assembly!")
        End
    End If
    
    'Get Assembly
    Dim oDoc As AssemblyDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oDef As AssemblyComponentDefinition
    Set oDef = oDoc.ComponentDefinition
    
    'Get Filename
    Dim fileName As String
    fileName = oDoc.DisplayName
    fileName = Replace(fileName, ".iam", "")

    'Create Path
    Dim path As String
    path = oDoc.FullFileName
    path = Left(path, Len(path) - Len(oDoc.DisplayName))
    path = path & "simplified\"
    
    createFolder path
    
    'Create Part
    Dim oPartDoc As PartDocument
    Set oPartDoc = ThisApplication.Documents.Add(kPartDocumentObject, , True)
    Dim oPartDef As PartComponentDefinition
    Set oPartDef = oPartDoc.ComponentDefinition
    
    Dim oDerivedAssemblyDef As DerivedAssemblyDefinition
    Set oDerivedAssemblyDef = oPartDef.ReferenceComponents.DerivedAssemblyComponents.CreateDefinition(oDoc.FullDocumentName)
    
    oDerivedAssemblyDef.DeriveStyle = kDeriveAsSingleBodyNoSeams
    oDerivedAssemblyDef.IncludeAllTopLevelWorkFeatures = kDerivedIncludeAll
    oDerivedAssemblyDef.IncludeAllTopLevelSketches = kDerivedIncludeAll
    oDerivedAssemblyDef.IncludeAllTopLeveliMateDefinitions = kDerivedIncludeAll
    oDerivedAssemblyDef.IncludeAllTopLevelParameters = kDerivedIncludeAll
    oDerivedAssemblyDef.ReducedMemoryMode = True
    Call oDerivedAssemblyDef.SetHolePatchingOptions(kDerivedPatchNone)
    Call oDerivedAssemblyDef.SetRemoveByVisibilityOptions(kDerivedRemovePartsAndFaces, 0.1)
    
    Dim oDerivedAssembly As DerivedAssemblyComponent
    Set oDerivedAssembly = oPartDef.ReferenceComponents.DerivedAssemblyComponents.Add(oDerivedAssemblyDef)
   
    Dim iptPath As String: iptPath = path & "" & fileName & ".ipt"
    Dim stepPath As String: stepPath = path & "" & fileName & ".step"
   
    Dim oTranslator As New MyTranslator
    Call oTranslator.ExportToSTEP(oPartDoc, stepPath)
    
    oPartDoc.Close True
    Set oPartDoc = ThisApplication.Documents.Open(stepPath)
    Call oPartDoc.SaveAs(iptPath, True)
    oPartDoc.Close True
    deleteFile stepPath
    oDoc.Close
End Sub


'Create all folders of a path
Public Sub createFolder(ByVal path As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim splittedPath() As String
    Dim newPath As String
    Dim loopStarted As Boolean
    newPath = ""
    loopStarted = False
    splittedPath = Split(path, "\")
    
    For Each s In splittedPath
        If s <> "" Then
            If Not loopStarted Then
                loopStarted = True
                newPath = s
            Else
                newPath = newPath & "\" & s
            End If
            If Not (fso.folderExists(newPath)) Then
                MkDir newPath
            End If
        End If
    Next
End Sub


Public Sub deleteFile(ByVal path As String)
    Dim FSO As FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If fileExists(path) Then
        FSO.deleteFile path
    End If
End Sub
