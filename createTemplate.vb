Imports System.Drawing
Imports System.IO
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor

Public Class createTemplate
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()

    End Sub

    Protected Overrides Sub OnClick()

        Dim pFLayer As IFeatureLayer
        Dim pLayer As ILayer
        Dim pGeoFtrLayer As IGeoFeatureLayer
        Dim pMXDoc As IMxDocument = My.ArcMap.Application.Document
        Dim pMap As IMap = pMXDoc.FocusMap
        Dim pEnumLayer As IEnumLayer
        Dim pFeatureSel As IFeatureSelection
        Dim pSelectionSet As ISelectionSet
        Dim pdataset As IDataset
        Dim pWorkspace As IWorkspace

        pEnumLayer = pMap.Layers
        pLayer = pEnumLayer.Next

        Do Until pLayer Is Nothing
            If TypeOf pLayer Is IFeatureLayer And pLayer.Visible = True Then
                pFLayer = pLayer
                pdataset = pLayer
                pGeoFtrLayer = pFLayer
                pWorkspace = pdataset.Workspace

                'start editing
                'get a reference to the editor
                Dim uid As UID
                uid = New UIDClass()
                uid.Value = "esriEditor.Editor"
                Dim pEditor As IEditor3
                pEditor = CType(My.ArcMap.Application.FindExtensionByCLSID(uid), IEditor3)

                'check if editing already
                If pEditor.EditState = esriEditState.esriStateNotEditing Then
                    pEditor.StartEditing(pWorkspace)
                Else
                    pEditor.StopEditing(True)
                End If


                'set up template manager
                Dim pEditTempManager As IEditTemplateManager = GetEditTemplateManager(pLayer)
                Dim pEditTemplate As IEditTemplate = Nothing

                pEditor.RemoveAllTemplatesInLayer(pLayer)

                Dim pEditTempFactory As IEditTemplateFactory = New EditTemplateFactory
                'pEditTemplate = pEditTempFactory.Create(pLayer.Name, pLayer)

                'get selected features from layer
                pFeatureSel = pLayer
                'pSelectionSet = pFeatureSel.SelectionSet
                Dim pfeature As IFeature

                Dim pTempArray As IArray = New Array
                Dim pRender As IUniqueValueRenderer = New UniqueValueRenderer


                'set geofeaturelayer renderer to unique value
                'Dim pUID As UID = New UID
                'pUID.Value = "{683C994E-A17B-11D1-8816-080009EC732A}"
                'pGeoFtrLayer.RendererPropertyPageClassID = pUID


                pRender = TryCast(pGeoFtrLayer.Renderer, IUniqueValueRenderer)
                'pRender = pGeoFtrLayer.Renderer
                'pUID = pGeoFtrLayer.RendererPropertyPageClassID
                'If pFLayer.FeatureClass.FindField("SYM") <> -1 Then
                '    pRender.FieldCount = 1
                '    pRender.Field(0) = "SYM"
                'Else
                '    pLayer = pEnumLayer.Next
                '    Continue Do
                'End If

                Dim pSymbolField As String = pRender.Field(0).ToString

                For pp As Integer = 0 To pRender.ValueCount - 1
                    Dim newEditTemplate As IEditTemplate = pEditTempFactory.Create(pRender.Label(pRender.Value(pp)), pLayer)

                    'need code to select feature with SYM value and get defaults
                    Dim pQFilt As IQueryFilter = New QueryFilter
                    pQFilt.WhereClause = "SYM = " + "'" + pRender.Value(pp) + "'"
                    pSelectionSet = pFLayer.FeatureClass.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, Nothing)
                    Dim lID As Long
                    Dim pEnumIDs As IEnumIDs = pSelectionSet.IDs
                    lID = pEnumIDs.Next
                    Do Until lID = -1
                        pfeature = pFLayer.FeatureClass.GetFeature(lID)

                        'newEditTemplate.SetDefaultValue(pSymbolField, pRender.Value(pp), False)

                        If pFLayer.FeatureClass.FindField("FSUBTYPE") <> -1 Then
                            Dim intFSUBTYPEidx As Integer = pFLayer.FeatureClass.FindField("FSUBTYPE")
                            newEditTemplate.SetDefaultValue("FSUBTYPE", pfeature.Value(intFSUBTYPEidx), True)
                        End If

                        If pFLayer.FeatureClass.FindField("FTYPE") <> -1 Then
                            Dim intFTYPEidx As Integer = pFLayer.FeatureClass.FindField("FTYPE")
                            newEditTemplate.SetDefaultValue("FTYPE", pfeature.Value(intFTYPEidx), True)
                        End If

                        If pFLayer.FeatureClass.FindField("POS") <> -1 Then
                            Dim intPOSidx As Integer = pFLayer.FeatureClass.FindField("POS")
                            newEditTemplate.SetDefaultValue("POS", pfeature.Value(intPOSidx), True)
                        End If

                        If pFLayer.FeatureClass.FindField("NOTES") <> -1 Then
                            Dim intNOTESidx As Integer = pFLayer.FeatureClass.FindField("NOTES")
                            newEditTemplate.SetDefaultValue("NOTES", pfeature.Value(intNOTESidx), True)
                        End If

                        If pFLayer.FeatureClass.FindField("GMAP_ID") <> -1 Then
                            Dim intNOTESidx As Integer = pFLayer.FeatureClass.FindField("GMAP_ID")
                            newEditTemplate.SetDefaultValue("GMAP_ID", pfeature.Value(intNOTESidx), True)
                        End If

                        newEditTemplate.SetDefaultValue(pSymbolField, pRender.Value(pp), False)
                        
                        lID = pEnumIDs.Next
                    Loop

                    pTempArray.Add(newEditTemplate)
                Next



                'Dim i As Integer
                'For i = 0 To pSelectionSet.Count - 1

                'pEditTemplate = pEditTempFactory.Create(pLayer.Name, pLayer)
                'pEditTemplate.SetDefaultValue("NOTES", "TEST", True)


                'loop thru selected features and add new template


                'Dim lID As Long
                'Dim pEnumIDs As IEnumIDs = pSelectionSet.IDs
                'lID = pEnumIDs.Next
                'Do Until lID = -1
                '    pfeature = pFLayer.FeatureClass.GetFeature(lID)
                '    pEditTemplate = pEditTempFactory.Create(pfeature.Value(pFLayer.FeatureClass.FindField("FTYPE")), pLayer)
                '    'pEditTemplate.SetDefaultValue("NOTES", "TEST", True)
                '    pTempArray.Add(pEditTemplate)
                '    lID = pEnumIDs.Next
                'Loop


                'add template to map
                'Dim pTempArray As IArray = New Array
                'pTempArray.Add(pEditTemplate)
                'pEditor.AddTemplates(pTempArray)
                'Next
                pEditor.AddTemplates(pTempArray)


                'Dim idx As Integer
                'For idx = 0 To pEditTempManager.Count - 1
                '    MsgBox(pEditTempManager.EditTemplate(idx).Name)
                'Next



                'loop through selection set, acquire default values and add to template for layer
                ' pEditTemplate.SetDefaultValue("NOTES", "TEST", True)

                pEditor.StopEditing(True)
            End If
            pLayer = pEnumLayer.Next
        Loop


        My.ArcMap.Application.CurrentTool = Nothing
    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub

    Private Function GetEditTemplateManager(ByVal layer As ILayer) As IEditTemplateManager
        Dim layerExtensions As ILayerExtensions
        Dim editTemplateMgr As IEditTemplateManager = Nothing
        layerExtensions = layer
        'Find the EditTemplateManager extension.
        For j As Integer = 0 To layerExtensions.ExtensionCount - 1
            Dim extension As Object = layerExtensions.Extension(j)
            If TypeOf extension Is IEditTemplateManager Then
                editTemplateMgr = extension
                Exit For
            End If
        Next
        'Use EditTemplateManager to get information about templates.
        Return editTemplateMgr
    End Function


End Class
