Option Explicit

Public Sub Example_AddPointsAtCrossings()
  Call AddPoints("Base de coupe", "PROTECAB Tube de protection de câble", "Intersection coupe", "Tube interne coupe")
End Sub


Public Sub AddPoints(sLineLayer1 As String, sLineLayer2 As String, sPointLayer As String, sPointLayer2 As String)
' Purpose:    Adds points wherever the line features from the specified layers cross
' Requires:   must be editing and have target point layer set
' Optionally: have some features selected in the first line layer

On Error GoTo EH
  Dim pApp As IApplication
  Dim pMxDoc As IMxDocument
  Dim pMap As IMap
  Dim pID As New UID
  Dim pEditor As IEditor
  Dim pELayers As IEditLayers
  Dim pFLayerLine1 As IFeatureLayer
  Dim pFLayerLine2 As IFeatureLayer
  Dim pFLayerCross As IFeatureLayer
  Dim pFLayerCross2 As IFeatureLayer
  Dim pFCLine1 As IFeatureClass
  Dim pFCLine2 As IFeatureClass
  Dim pFCCross As IFeatureClass
  Dim pFCCross2 As IFeatureClass
  Dim pFSel1 As IFeatureSelection
  Dim pFCursor1 As IFeatureCursor
  Dim pFCursor2 As IFeatureCursor
  Dim pFeature1 As IFeature
  Dim pFeature2 As IFeature
  Dim pNewFeature As IFeature
  Dim pNewFeature2 As IFeature
  Dim lCount As Long
  Dim lTotal As Long
  Dim pRowSubtypes As IRowSubtypes
  Dim pSubtypes As ISubtypes
  Dim lSubCode As Long
  Dim bHasSubtypes As Boolean
  Dim pCurve1 As ICurve
  Dim pCurve2 As ICurve
  Dim pTopoOp1 As ITopologicalOperator
  Dim pEnv1 As IEnvelope
  Dim pSFilter As ISpatialFilter
  Dim pPoint As IPoint
  Dim pMPoint As IMultipoint
  Dim pGeoCol As IGeometryCollection
  Dim lGeoTotal As Long
  Dim lGeoCount As Integer
  
  Dim lIDCoupe As Long
  Dim lIDTube As Long
  Dim pQueryFilter As IQueryFilter
  Dim pFSel2 As IFeatureSelection
  Dim pFSel3 As IFeatureSelection
  Dim pDeleteSet As ISet
  Dim pFeatureEdit As IFeatureEdit
  
  Set pApp = Application
  Set pMxDoc = pApp.Document
  Set pMap = pMxDoc.FocusMap
  
  ' Verify that there are at least 2 layers in the table on contents
  If pMap.LayerCount < 2 Then
    MsgBox "Must have at least two layers in your map."
    Exit Sub
  End If
  
  'Find the two line layers by name
  Set pFLayerLine1 = FindFLayerByName(pMap, sLineLayer1)
  Set pFLayerLine2 = FindFLayerByName(pMap, sLineLayer2)
  
  'Verify layers exist
  If pFLayerLine1 Is Nothing Then
    MsgBox sLineLayer1 & " layer not found."
    Exit Sub
  End If
  If pFLayerLine2 Is Nothing Then
    MsgBox sLineLayer2 & " layer not found."
    Exit Sub
  End If
  
  Set pFCLine1 = pFLayerLine1.FeatureClass
  Set pFCLine2 = pFLayerLine2.FeatureClass
  
  'Verify that it is a correct type of geometry
  If pFCLine1.ShapeType <> esriGeometryPolyline And pFCLine1.ShapeType <> esriGeometryLine Then
    MsgBox sLineLayer1 & " layer must be a line or polyline layer."
    Exit Sub
  End If
  If pFCLine2.ShapeType <> esriGeometryPolyline And pFCLine2.ShapeType <> esriGeometryLine Then
    MsgBox sLineLayer2 & " layer must be a line or polyline layer."
    Exit Sub
  End If
  
'  'Verify that the target is a point layer
'  Set pELayers = pEditor
'  If pELayers.CurrentLayer.FeatureClass.ShapeType = esriGeometryMultipoint Then
'    MsgBox "This edit target is a multipoint layer.  Please use a point layer." & vbNewLine & "Convert using ""Multipart To Singlepart"" GP tool if needed."
'    Exit Sub
'  End If
'
'  If pELayers.CurrentLayer.FeatureClass.ShapeType <> esriGeometryPoint Then
'    MsgBox "Edit target must be a point layer (i.e. crossing points)."
'    Exit Sub
'  End If
    
  'Get the target point layer
  'Set pFLayerCross = pELayers.CurrentLayer
  Set pFLayerCross = FindFLayerByName(pMap, sPointLayer)
  Set pFCCross = pFLayerCross.FeatureClass
  
  Set pFLayerCross2 = FindFLayerByName(pMap, sPointLayer2)
  Set pFCCross2 = pFLayerCross2.FeatureClass
 
  'Get current target subtype
  If TypeOf pFCCross Is ISubtypes Then
    Set pSubtypes = pFCCross
    If pSubtypes.HasSubtype Then
      bHasSubtypes = True
      lSubCode = pELayers.CurrentSubtype
    Else
      bHasSubtypes = False
    End If
  End If
  
  'Verify that we are editing
  pID = "esriEditor.Editor"
  Set pEditor = pApp.FindExtensionByCLSID(pID)
  If Not (pEditor.EditState = esriStateEditing) Then
    'MsgBox "Must be editing."
    Call StartEditing(pFLayerCross)
    'Exit Sub
  End If
  
  
  'Update Message bar
  pApp.StatusBar.Message(0) = "Adding " & pFLayerCross.Name & " points..."
  
  'Start edit operation (for undo)
  pEditor.StartOperation
  
  'Now that an edit operation has been started, use a different error handler
  'in order to abort this operation if a problem occurs
  On Error GoTo EH2
  
  'If any features in the first layer are selected, use them only
  'Otherwise diplay error message
  Set pFSel1 = pFLayerLine1
  If pFSel1.SelectionSet.Count > 0 Then
    pFSel1.SelectionSet.Search Nothing, False, pFCursor1
    lTotal = pFSel1.SelectionSet.Count
  Else
    MsgBox "Une coupe au moins doit être sélectionnée"
    pEditor.StopOperation ("Add Points")
    Exit Sub
  End If
  
  'Step through each feature in layer1
  lCount = 1
  Set pFeature1 = pFCursor1.NextFeature
  Do While Not pFeature1 Is Nothing
    
    'Update status bar
    lCount = lCount + 1
    pApp.StatusBar.Message(0) = "Processing " & pFLayerLine1.Name & " lines ..." & Str(lCount) & " of " & Str(lTotal)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    'Select old crossing points
    
    'Get OBJECTID of pFeature1
    lIDCoupe = pFeature1.Value(pFeature1.Class.FindField("OBJECTID"))
    
    'Remove Points
    Set pQueryFilter = New QueryFilter
    pQueryFilter.whereClause = "QS_BASE_REF = " & lIDCoupe
    
    Set pFSel2 = pFLayerCross
    Set pFSel3 = pFLayerCross2
    
    Set pDeleteSet = New esriSystem.Set
    pMap.ClearSelection
    'MsgBox lIDCoupe
    pFSel2.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    pFSel3.SelectFeatures pQueryFilter, esriSelectionResultNew, False
       
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Get needed references to this feature from layer1
    Set pCurve1 = pFeature1.Shape
    Set pTopoOp1 = pCurve1
    Set pEnv1 = pCurve1.Envelope
    
    'Create a spatial filter for layer2 to find any potentially crossing lines
    Set pSFilter = New SpatialFilter
    Set pSFilter.Geometry = pEnv1
    pSFilter.GeometryField = pFCLine2.ShapeFieldName
    pSFilter.SpatialRel = esriSpatialRelIntersects
    Set pFCursor2 = pFCLine2.Search(pSFilter, False)
    
    ' Step through each feature in layer2 that crosses the envelope of the
    ' current feature we are processing from layer1
    Set pFeature2 = pFCursor2.NextFeature
    Do While Not pFeature2 Is Nothing
      
      'Get OBJECTID of pFeature2
      lIDTube = pFeature2.Value(pFeature2.Class.FindField("OBJECTID"))
      
      'MsgBox lIDTube
          
      'Get the geometry for this feature from layer2
      If pFLayerLine1 Is pFLayerLine2 Then
        Set pCurve2 = pFeature2.ShapeCopy
      Else
        Set pCurve2 = pFeature2.Shape
      End If
      
      'Find all intersecting points (returned as multipoint)
      Set pMPoint = pTopoOp1.Intersect(pCurve2, esriGeometry0Dimension)
      If Not pMPoint Is Nothing Then
        If Not pMPoint.IsEmpty Then
          Set pGeoCol = pMPoint
          
          'Step through each point in the multipoint (often just one)
          lGeoTotal = pGeoCol.GeometryCount
          For lGeoCount = 0 To lGeoTotal - 1
          
            'Get the point
            Set pPoint = pGeoCol.Geometry(lGeoCount)
            
            'Create the new feature and set it's geometry
            Set pNewFeature = pFCCross.CreateFeature
            Set pNewFeature.Shape = pPoint
            
            'Set attributes
            pNewFeature.Value(pNewFeature.Class.FindField("QS_BASE_REF")) = lIDCoupe
            pNewFeature.Value(pNewFeature.Class.FindField("ROHR_REF")) = lIDTube
            pNewFeature.Value(pNewFeature.Class.FindField("DIAMETRE")) = pFeature2.Value(pFeature2.Class.FindField("U_DIAMETRE"))
            pNewFeature.Value(pNewFeature.Class.FindField("DIAMETRE_MM")) = pFeature2.Value(pFeature2.Class.FindField("U_DIAMETRE_MM"))
            pNewFeature.Value(pNewFeature.Class.FindField("KABELSCHUTZ")) = pFeature2.Value(pFeature2.Class.FindField("KABELSCHUTZ"))
          
            'If needed, set the subtype and default values
            If bHasSubtypes Then
              Set pRowSubtypes = pNewFeature
              pRowSubtypes.SubtypeCode = lSubCode
              pRowSubtypes.InitDefaultValues
            End If
            'Save the new feature
            pNewFeature.Store
            
          Next lGeoCount
        End If
      End If

      Set pFeature2 = pFCursor2.NextFeature
    Loop
    
    ''''''''''''''''''''''''''''''''''''''''''''
    'Remove old crossing points
    If pFSel2.SelectionSet.Count > 0 Then
        pFSel2.SelectionSet.Search Nothing, False, pFCursor2
    
        Set pFeature2 = pFCursor2.NextFeature
        Do While Not pFeature2 Is Nothing
            'MsgBox "aaa"
            Set pFeatureEdit = pFeature2
            pDeleteSet.Add pFeature2
            Set pFeature2 = pFCursor2.NextFeature
        Loop
        pFeatureEdit.DeleteSet pDeleteSet
    End If
    
    'Remove old crossing points (tubes internes)
    If pFSel3.SelectionSet.Count > 0 Then
        pFSel3.SelectionSet.Search Nothing, False, pFCursor2
    
        Set pFeature2 = pFCursor2.NextFeature
        Do While Not pFeature2 Is Nothing
            'MsgBox "aaa"
            Set pFeatureEdit = pFeature2
            pDeleteSet.Add pFeature2
            Set pFeature2 = pFCursor2.NextFeature
        Loop
        pFeatureEdit.DeleteSet pDeleteSet
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    
    Set pFeature1 = pFCursor1.NextFeature
  Loop

  'Stop feature editing
  pEditor.StopOperation ("Add Points")
  
  'Clear all feature selections
  pMap.ClearSelection
  'pMap.SelectFeature(pFLayerLine1, pFSel1)
  
  'Redraw the map so you'll see the new lines
  pMxDoc.UpdateContents
  pMxDoc.ActiveView.Refresh
 
  'MsgBox "Auto Add Points is complete."
  pApp.StatusBar.Message(0) = "Add Points is complete."
  
  Exit Sub
EH:
   
  MsgBox Err.Number & "  " & Err.Description
  Exit Sub
EH2:
  pEditor.AbortOperation
  MsgBox Err.Number & "  " & Err.Description
  Exit Sub
End Sub

Public Sub StartEditing(pFLayerTarget As IFeatureLayer)

  Dim pEditor As IEditor
  Dim pID As New UID
  Dim pDataset As IDataset
  Dim pMap As IMap
  Dim pMxDoc As IMxDocument
  Dim LayerCount As Integer

  Set pMxDoc = Application.Document
  Set pMap = pMxDoc.FocusMap
  pID = "esriEditor.Editor"
  Set pEditor = Application.FindExtensionByCLSID(pID)

  If pEditor.EditState = esriStateEditing Then Exit Sub

  'Start editing the workspace of the first featurelayer you find
  'For LayerCount = 0 To pMap.LayerCount - 1
    'If TypeOf pMap.Layer(LayerCount) Is IFeatureLayer Then
      'Set pFeatureLayer = pMap.Layer(LayerCount)
      Set pDataset = pFLayerTarget.FeatureClass
      pEditor.StartEditing pDataset.Workspace
      'Exit For
    'End If
  'Next LayerCount

End Sub


Public Function FindFLayerByName(pMap As IMap, sLayerName As String) As IFeatureLayer
  'This function will return only feature layers.
  'It can find feature layers within groups.
  
  Dim pEnumLayer As IEnumLayer
  Dim pCompositeLayer As ICompositeLayer
  Dim i As Integer
  
  Set pEnumLayer = pMap.Layers
  pEnumLayer.Reset

  Dim pLayer As ILayer
  Set pLayer = pEnumLayer.Next

  Do While Not pLayer Is Nothing
    If TypeOf pLayer Is ICompositeLayer Then
      Set pCompositeLayer = pLayer
      For i = 0 To pCompositeLayer.Count - 1
        With pCompositeLayer
        If .Layer(i).Name = sLayerName Then
          If TypeOf .Layer(i) Is IFeatureLayer Then
            Set FindFLayerByName = pCompositeLayer.Layer(i)
            Exit Function
          End If
        End If
        End With
      Next i
    ElseIf pLayer.Name = sLayerName And TypeOf pLayer Is IFeatureLayer Then
      Set FindFLayerByName = pLayer
      Exit Function
    End If
    Set pLayer = pEnumLayer.Next
  Loop

End Function





