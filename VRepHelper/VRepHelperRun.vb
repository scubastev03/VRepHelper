Imports Inventor
Imports System.Data
Imports System.Xml
Public Class VRepHelperRun
    Private oDataSet As New DataSet()
    Private oTimerStart As DateTime ' = DateAndTime.Now()
    Private oInvApp As Application = VRepHelper.StandardAddInServer.oInvApp
    Private oDoc As Document = oInvApp.ActiveDocument
    Private oAsmCmpDef As AssemblyComponentDefinition = oDoc.ComponentDefinition
    Private oCmpOccs As ComponentOccurrences = oAsmCmpDef.Occurrences
    Private oViewReps As DesignViewRepresentations = oAsmCmpDef.RepresentationsManager.DesignViewRepresentations
    Private oViewRepActive As DesignViewRepresentation = oAsmCmpDef.RepresentationsManager.ActiveDesignViewRepresentation()
    Private oViewRep As DesignViewRepresentation
    Private oPropSets As PropertySets
    Private oLODReps As LevelOfDetailRepresentations = oAsmCmpDef.RepresentationsManager.LevelOfDetailRepresentations
    Private oLODRep As LevelOfDetailRepresentation
    Private oObjCol1 As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection
    Private oObjCol2 As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection
    Private oFileCount As Integer = 0

    Private Sub VRepHelperRun_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            Try
                oDataSet.ReadXml("VRepHelperDb.xml")
            Catch

            End Try
        Else
            MsgBox("Must be an asssembly file.")
            Exit Sub
        End If

    End Sub


    Public Sub Process()
        Try

            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then 'Check If File Is An Assembly

                'Try
                oFileCount = oDoc.AllReferencedDocuments.Count()
                'MsgBox(oFileCount.ToString)
                'Catch
                'End Try

                'Set Level Of Detail to Master
                'oLODReps.Item("Master").Activate()

                'For Each ORepDict In oRepDicts 'Create View Representations If Not Exist
                '    Try
                '        oViewReps.Item(ORepDict.Key).Activate()
                '    Catch
                '        oViewRep = oViewReps.Add(ORepDict.Key)
                '    End Try
                'Next

                'Run Through Each View Representation and Set Visibility

                For Each oViewRep In oViewReps

                    If oViewRep.DesignViewType() = DesignViewTypeEnum.kPublicDesignViewType Then
                        'Activate View Representation
                        oViewRep.Activate()
                        oInvApp.StatusBarText() = "Setting """ & oViewRep.Name & """ to active View Representation"

                        'Clear Object Collections
                        oObjCol1.Clear()
                        oObjCol2.Clear()

                        'Get Leaf Occurences for Current Assembly Document Definition
                        For Each oCmpOcc As ComponentOccurrence In oAsmCmpDef.Occurrences

                            If oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kAssemblyDocumentObject Then
                                If oCmpOcc.BOMStructure() = BOMStructureEnum.kPhantomBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                    Try
                                        oCmpOcc.SetDesignViewRepresentation(oViewRep.Name, Nothing, True)
                                    Catch
                                        Continue For
                                    End Try
                                Else
                                    Continue For
                                End If

                            ElseIf oCmpOcc.DefinitionDocumentType() = DocumentTypeEnum.kPartDocumentObject Then
                                If oCmpOcc.BOMStructure() = BOMStructureEnum.kNormalBOMStructure Or BOMStructureEnum.kPurchasedBOMStructure Or BOMStructureEnum.kReferenceBOMStructure Then
                                    oPropSets = oCmpOcc.Definition.Document.PropertySets

                                    If Not oPropSets.PropertySetExists("Inventor Summary Information") Then
                                        Continue For
                                    Else

                                        MsgBox(oPropSets.Item("Inventor Summary Information").Item("Title").Value)

                                    End If

                                End If
                            Else
                                Continue For
                            End If

                        Next

                        If oObjCol1.Count() <> 0 Then
                            Try
                                oViewRep.SetVisibilityOfOccurrences(oObjCol1, True)
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If

                        If oObjCol2.Count() <> 0 Then
                            Try
                                oViewRep.SetVisibilityOfOccurrences(oObjCol2, False)
                            Catch ex As Exception
                                MsgBox(ex.ToString)
                            End Try
                        End If



                        oInvApp.StatusBarText() = "Applied Component Visibiltiy to """ & oViewRep.Name & """"

                        Try
                            oLODReps.Item(oViewRep.Name).Delete()
                        Catch
                        End Try
                        If oFileCount > 100 Then
                            oViewReps.Item(oViewRep.Name).CopyToLevelOfDetail(oViewRep.Name)
                        End If

                    End If
                Next

            End If

            'oViewRep = oViewReps.Item("Default")
            'oViewRep.Activate()
            'oViewRepActive.Activate()

            Dim TimeSpent As System.TimeSpan = Now.Subtract(oTimerStart)

            oDoc.Save()

            MsgBox("Setting Design View & LOD Representation Process Completed." & vbCr & "Process completed in " & TimeSpent.Seconds & " seconds.",, "Design View & LOD Representations")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub


End Class