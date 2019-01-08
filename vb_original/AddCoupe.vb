Option Explicit
Public Sub Example_AddCoupe()
  Call AddPointsCoupe("Base de coupe", "Surface de coupe", "Intersection coupe", "CÂBLE Conduite de câble", "eler_rohr_kabel", "Câble coupe", "Tube interne coupe")
End Sub



Public Sub AddPointsCoupe(sLineLayer1 As String, sAreaLayer2 As String, sPointLayer3 As String, sLineLayer4 As String, sTable5 As String, sPointLayer6 As String, sPointLayer7 As String)
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
  Dim pFLayerArea2 As IFeatureLayer
  Dim pFLayerPoint3 As IFeatureLayer
  Dim pFLayerLine4 As IFeatureLayer
  Dim pFLayerPoint6 As IFeatureLayer
  Dim pFLayerPoint7 As IFeatureLayer

  Dim pFCLine1 As IFeatureClass
  Dim pFCArea2 As IFeatureClass
  Dim pFCPoint3 As IFeatureClass
  Dim pFCLine4 As IFeatureClass
  Dim pTTable5 As ITable
  Dim pFCPoint6 As IFeatureClass
  Dim pFCPoint7 As IFeatureClass
  
  Dim pFSel1 As IFeatureSelection
  Dim pFCursor1 As IFeatureCursor
  Dim pFCursor2 As IFeatureCursor
  Dim pFCursor4 As IFeatureCursor
  Dim pFCursor6 As IFeatureCursor
  Dim pFeature1 As IFeature
  Dim pFeature2 As IFeature
  Dim pFeature3 As IFeature
  Dim pFeature4 As IFeature
  Dim pDeleteSet As ISet
  Dim pFeatureEdit As IFeatureEdit
   
  Dim pNewFeature As IFeature
  Dim pNewFeature2 As IFeature
  Dim pNewFeature6 As IFeature
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
  Dim lFactor As Long
  Dim pQueryFilter As IQueryFilter
  Dim pQueryFilter2 As IQueryFilter2
  Dim pFSel2 As IFeatureSelection
  Dim pFSel3 As IFeatureSelection
  Dim pFSel4 As IFeatureSelection
  
  Dim Pi As Double
 
  Set pApp = Application
  Set pMxDoc = pApp.Document
  Set pMap = pMxDoc.FocusMap
  
  lFactor = 3
  Pi = 4 * Atn(1)
  
  'Find the layers by name
  Set pFLayerLine1 = FindFLayerByName(pMap, sLineLayer1)
  Set pFLayerArea2 = FindFLayerByName(pMap, sAreaLayer2)
  Set pFLayerPoint3 = FindFLayerByName(pMap, sPointLayer3)
  Set pFLayerLine4 = FindFLayerByName(pMap, sLineLayer4)
  Set pFLayerPoint6 = FindFLayerByName(pMap, sPointLayer6)
  Set pFLayerPoint7 = FindFLayerByName(pMap, sPointLayer7)

  Set pTTable5 = TrouveTable(sTable5)
  
  'Verify layers exist
  If pFLayerLine1 Is Nothing Then
    MsgBox sLineLayer1 & " layer not found."
    Exit Sub
  End If
  If pFLayerArea2 Is Nothing Then
    MsgBox sAreaLayer2 & " layer not found."
    Exit Sub
  End If
  If pFLayerPoint3 Is Nothing Then
    MsgBox sPointLayer3 & " layer not found."
    Exit Sub
  End If
  If pFLayerLine4 Is Nothing Then
    MsgBox sLineLayer4 & " layer not found."
    Exit Sub
  End If
  If pTTable5 Is Nothing Then
    MsgBox sTable5 & " layer not found."
    Exit Sub
  End If
  If pFLayerPoint6 Is Nothing Then
    MsgBox sPointLayer6 & " layer not found."
    Exit Sub
  End If
  If pFLayerPoint7 Is Nothing Then
    MsgBox sPointLayer7 & " layer not found."
    Exit Sub
  End If
  
  Set pFCLine1 = pFLayerLine1.FeatureClass
  Set pFCArea2 = pFLayerArea2.FeatureClass
  Set pFCPoint3 = pFLayerPoint3.FeatureClass
  Set pFCLine4 = pFLayerLine4.FeatureClass
  Set pFCPoint6 = pFLayerPoint6.FeatureClass
  Set pFCPoint7 = pFLayerPoint7.FeatureClass
  
  'Verify that it is a correct type of geometry
  If pFCLine1.ShapeType <> esriGeometryPolyline And pFCLine1.ShapeType <> esriGeometryLine Then
    MsgBox sLineLayer1 & " layer must be a line or polyline layer."
    Exit Sub
  End If
  If pFCArea2.ShapeType <> esriGeometryPolygon Then
    MsgBox sAreaLayer2 & " layer must be a polygon layer."
    Exit Sub
  End If
  If pFCPoint3.ShapeType <> esriGeometryPoint Then
    MsgBox sPointLayer3 & " layer must be a point layer."
    Exit Sub
  End If
  If pFCLine4.ShapeType <> esriGeometryPolyline And pFCLine4.ShapeType <> esriGeometryLine Then
    MsgBox sLineLayer4 & " layer must be a line or polyline layer."
    Exit Sub
  End If
  If pFCPoint6.ShapeType <> esriGeometryPoint Then
    MsgBox sPointLayer6 & " layer must be a point layer."
    Exit Sub
  End If
  If pFCPoint7.ShapeType <> esriGeometryPoint Then
    MsgBox sPointLayer7 & " layer must be a point layer."
    Exit Sub
  End If
 
'  'Get current target subtype
'  If TypeOf pFCCross Is ISubtypes Then
'    Set pSubtypes = pFCCross
'    If pSubtypes.HasSubtype Then
'      bHasSubtypes = True
'      lSubCode = pELayers.CurrentSubtype
'    Else
'      bHasSubtypes = False
'    End If
'  End If
  
  'Verify that we are editing
  pID = "esriEditor.Editor"
  Set pEditor = pApp.FindExtensionByCLSID(pID)
  If Not (pEditor.EditState = esriStateEditing) Then
    'MsgBox "Must be editing."
    Call StartEditing(pFLayerArea2)
    'Exit Sub
  End If
  
  'Update Message bar
  'pApp.StatusBar.Message(0) = "Adding " & pFLayerPoint3.Name & " points..."
  
  'Start edit operation (for undo)
  pEditor.StartOperation
  
  'Now that an edit operation has been started, use a different error handler
  'in order to abort this operation if a problem occurs
  On Error GoTo EH2
  
  'If any features in the first layer are selected, use them only
  'Otherwise use all features from the first layer
  Set pFSel1 = pFLayerLine1
  If pFSel1.SelectionSet.Count > 0 Then
    pFSel1.SelectionSet.Search Nothing, False, pFCursor1
    lTotal = pFSel1.SelectionSet.Count
  Else
    MsgBox "Une coupe au moins doit être sélectionnée"
    pEditor.StopOperation ("Add Points")
    Exit Sub
  End If
 
  'Step through each feature in layer1 (Coupe)
  lCount = 1
  Set pFeature1 = pFCursor1.NextFeature
  'Parcours de chaque coupe
  Do While Not pFeature1 Is Nothing
    
    'Get OBJECTID of pFeature1
    lIDCoupe = pFeature1.Value(pFeature1.Class.FindField("OBJECTID"))
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Remove Surface
    Set pQueryFilter = New QueryFilter
    pQueryFilter.whereClause = "QS_REF = " & lIDCoupe
    
    Set pFSel2 = pFLayerArea2
    Set pFSel3 = pFLayerPoint6
    
    Set pDeleteSet = New esriSystem.Set
    pMap.ClearSelection
    pFSel2.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    pFSel3.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    
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
    
    If pFSel3.SelectionSet.Count > 0 Then
        pFSel3.SelectionSet.Search Nothing, False, pFCursor2
        
        Set pFeature2 = pFCursor2.NextFeature
        Do While Not pFeature2 Is Nothing
            'MsgBox "bbb"
            Set pFeatureEdit = pFeature2
            pDeleteSet.Add pFeature2
            Set pFeature2 = pFCursor2.NextFeature
        Loop
        pFeatureEdit.DeleteSet pDeleteSet
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Select Points
    Set pQueryFilter = New QueryFilter
    pQueryFilter.whereClause = "QS_BASE_REF = " & lIDCoupe
    
    Set pFSel2 = pFLayerPoint3
    
    pMap.ClearSelection
    
    pFSel2.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    
    If pFSel2.SelectionSet.Count > 0 Then
        pFSel2.SelectionSet.Search Nothing, False, pFCursor2
    Else
        MsgBox "Aucun point de coupe, coupe: " & lIDCoupe
        pEditor.StopOperation ("Add Points")
        Exit Sub
    End If
    
    'Parcours des points pour obtenir le gabarit de la coupe
    Dim lAbsPointMax As Double
    Dim lAbsPointMin As Double
    Dim lProfPointMax As Double
    Dim lProfPointMin As Double
    
    lAbsPointMax = 0
    lAbsPointMin = 1000
    lProfPointMax = 0
    lProfPointMin = 1000
    
    Set pFeature2 = pFCursor2.NextFeature
    'Parcours de chaque tube
    Do While Not pFeature2 Is Nothing
        
        Set pPoint = pFeature2.Shape
        If pFeature2.Value(pFeature2.Class.FindField("ABSCISSE")) + pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000 > lAbsPointMax Then
            lAbsPointMax = pFeature2.Value(pFeature2.Class.FindField("ABSCISSE")) + pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000
        End If
        
        If pFeature2.Value(pFeature2.Class.FindField("ABSCISSE")) - pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000 < lAbsPointMin Then
            lAbsPointMin = pFeature2.Value(pFeature2.Class.FindField("ABSCISSE")) - pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000
        End If
        
        If Not IsNull(pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR"))) Then
            If pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR")) + pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 1000 > lProfPointMax Then
                lProfPointMax = pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR")) + pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 1000
            End If
            
            If pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR")) < lProfPointMin Then
                lProfPointMin = pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR"))
            End If
            Else
                If Not IsNull(pFeature2.Value(pFeature2.Class.FindField("TUBEEXT"))) Then
                    MsgBox "Tube interne détecté"
                Else
                    MsgBox "Il manque des profondeurs de protection de câbles"
                    pEditor.StopOperation ("Add Points")
                    Exit Sub
                End If
            End If
        
        Set pFeature2 = pFCursor2.NextFeature
    Loop
    
    Dim pPointC0 As IPoint
    Dim pPointC1 As IPoint
    Dim pPointC2 As IPoint
    Dim pPointC3 As IPoint
    Dim pPointC4 As IPoint
    Dim pCurveC As ICurve
    Dim pLineC As ILine
    
    Set pCurveC = pFeature1.Shape
    Set pPointC0 = pCurveC.FromPoint
    Set pPointC1 = pCurveC.ToPoint
    Set pLineC = New Line
    
    pLineC.PutCoords pPointC1, pPointC0
    Dim lAngleC, lAngleCdeg, lLengthC, lProfC, lLengthLineC As Double
    lAngleC = pLineC.Angle
    lAngleCdeg = (180 * lAngleC) / Pi
    lLengthLineC = pLineC.Length

    lLengthC = (lAbsPointMax - lAbsPointMin + 0.1) * lFactor
    lProfC = (lProfPointMax - lProfPointMin + 0.1) * lFactor
    
    'Calcul de la coupe
    Set pPointC2 = New Point
    Set pPointC3 = New Point
    Set pPointC4 = New Point
    
    pPointC2.X = -lLengthC * Cos(lAngleC) + pPointC1.X
    pPointC2.Y = -lLengthC * Sin(lAngleC) + pPointC1.Y
    
    If Cos(lAngleC) > 0 Then
        pPointC3.X = -lProfC * Sin(lAngleC) + pPointC2.X
        pPointC3.Y = lProfC * Cos(lAngleC) + pPointC2.Y
        
        pPointC4.X = -lProfC * Sin(lAngleC) + pPointC1.X
        pPointC4.Y = lProfC * Cos(lAngleC) + pPointC1.Y
    Else
        pPointC3.X = lProfC * Sin(lAngleC) + pPointC2.X
        pPointC3.Y = -lProfC * Cos(lAngleC) + pPointC2.Y
        
        pPointC4.X = lProfC * Sin(lAngleC) + pPointC1.X
        pPointC4.Y = -lProfC * Cos(lAngleC) + pPointC1.Y
    End If
    
    Dim pSegCollC As ISegmentCollection
    Set pSegCollC = New Ring
    Dim pLine As ILine
    Set pLine = New Line
    pLine.PutCoords pPointC1, pPointC4
    pSegCollC.AddSegment pLine
    Set pLine = New Line
    pLine.PutCoords pPointC4, pPointC3
    pSegCollC.AddSegment pLine
    Set pLine = New Line
    pLine.PutCoords pPointC3, pPointC2
    pSegCollC.AddSegment pLine
    Dim pRingC As IRing
    Set pRingC = pSegCollC
    pRingC.Close
    
    Dim pPolygonC As IGeometryCollection
    Set pPolygonC = New Polygon
    pPolygonC.AddGeometry pRingC
    
    Set pNewFeature = pFCArea2.CreateFeature
    Set pNewFeature.Shape = pPolygonC
    
    'Ajout attibuts à surface de coupe
    pNewFeature.Value(pNewFeature.Class.FindField("PROF_MAX")) = -lProfPointMax
    pNewFeature.Value(pNewFeature.Class.FindField("PROF_MIN")) = -lProfPointMin
    pNewFeature.Value(pNewFeature.Class.FindField("QS_REF")) = lIDCoupe
    
    pNewFeature.Store
            
    'Select Points
    Set pQueryFilter = New QueryFilter
    
    pQueryFilter.whereClause = "QS_BASE_REF = " & lIDCoupe & "AND TUBEEXT is Null"
    
    Set pFSel2 = pFLayerPoint3
    
    pMap.ClearSelection
    
    pFSel2.SelectFeatures pQueryFilter, esriSelectionResultNew, False
    
    If pFSel2.SelectionSet.Count > 0 Then
        pFSel2.SelectionSet.Search Nothing, False, pFCursor2
    Else
        MsgBox "Aucun point de coupe"
        pEditor.StopOperation ("Add Points")
        Exit Sub
    End If
    
    Set pFeature2 = pFCursor2.NextFeature
    'Parcours de chaque tube
    Do While Not pFeature2 Is Nothing
        
        'Dim pPointC As IPoint
        Dim lAbsPoint As Double
        Dim lProfPoint As Double
        Dim lRohrRef As Long
        
        Set pPoint = pFeature2.Shape
        lRohrRef = pFeature2.Value(pFeature2.Class.FindField("ROHR_REF"))
        lAbsPoint = pFeature2.Value(pFeature2.Class.FindField("ABSCISSE"))
        
        'Tubes dans tube courant--------------------------------------------------------------------------------------------------------------------------------------
        Set pQueryFilter2 = New QueryFilter
        pQueryFilter2.whereClause = "QS_BASE_REF = " & lIDCoupe & "AND TUBEEXT = " & lRohrRef
        Set pFSel4 = pFLayerPoint7
        pFSel4.SelectFeatures pQueryFilter2, esriSelectionResultNew, False
        '-----------------------------------------------------------------------------------------------------------------------------------------------------
        
        If IsNull(pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR"))) Then
            lProfPoint = 0
        Else
            lProfPoint = pFeature2.Value(pFeature2.Class.FindField("PROFONDEUR")) + pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000
        End If
        
        If Cos(lAngleC) > 0 Then
            pPoint.X = (-(lAbsPoint - lAbsPointMin + (lLengthC / lFactor - (lAbsPointMax - lAbsPointMin)) / 2) * Cos(lAngleC) - (lProfC / lFactor - (lProfPoint - lProfPointMin + (lProfC / lFactor - (lProfPointMax - lProfPointMin)) / 2)) * Cos(Pi / 2 - lAngleC)) * lFactor + pPointC1.X
            pPoint.Y = (-(lAbsPoint - lAbsPointMin + (lLengthC / lFactor - (lAbsPointMax - lAbsPointMin)) / 2) * Sin(lAngleC) + (lProfC / lFactor - (lProfPoint - lProfPointMin + (lProfC / lFactor - (lProfPointMax - lProfPointMin)) / 2)) * Sin(Pi / 2 - lAngleC)) * lFactor + pPointC1.Y
        Else
            pPoint.X = (-(lAbsPoint - lAbsPointMin + (lLengthC / lFactor - (lAbsPointMax - lAbsPointMin)) / 2) * Cos(lAngleC) + (lProfC / lFactor - (lProfPoint - lProfPointMin + (lProfC / lFactor - (lProfPointMax - lProfPointMin)) / 2)) * Cos(Pi / 2 - lAngleC)) * lFactor + pPointC1.X
            pPoint.Y = (-(lAbsPoint - lAbsPointMin + (lLengthC / lFactor - (lAbsPointMax - lAbsPointMin)) / 2) * Sin(lAngleC) - (lProfC / lFactor - (lProfPoint - lProfPointMin + (lProfC / lFactor - (lProfPointMax - lProfPointMin)) / 2)) * Sin(Pi / 2 - lAngleC)) * lFactor + pPointC1.Y
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim MaTableCursor As ICursor 'Le curseur sur le résultat de la query
        Dim MaTableRow As IRow 'Un enregistrement dans la table
        Dim MaTableQF As IQueryFilter2 'Le query filter
        Dim lKabelRef As Long
        Dim lKabelSpannung As Long
        Dim lKabelCount As Long
        Dim lKabelTotal As Long
                
        lKabelCount = 0
        
        If pFSel4.SelectionSet.Count > 0 Then
            MsgBox pFSel4.SelectionSet.Count & " tube(s) interne(s)"
            pFSel4.SelectionSet.Search Nothing, False, pFCursor4
            lKabelTotal = pFSel4.SelectionSet.Count
            Set pFeature4 = pFCursor4.NextFeature
        Else
            'MsgBox "Pas de tube interne"
            lKabelTotal = 0
        End If
        
        Set MaTableQF = New QueryFilter
        
        'On set le QueryFilter
        MaTableQF.whereClause = "[ROHR_REF] =" & lRohrRef
        'On applique le QueryFilter
        Set MaTableCursor = pTTable5.Search(MaTableQF, False)
        
        'Décompte des câbles
        Set MaTableRow = MaTableCursor.NextRow
        Do While Not MaTableRow Is Nothing
            lKabelTotal = lKabelTotal + 1
            Set MaTableRow = MaTableCursor.NextRow
        Loop
        
        'Récupération des attributs du câble, via les relations
        Set MaTableCursor = pTTable5.Search(MaTableQF, False)
        Set MaTableRow = MaTableCursor.NextRow
        'Parcours de chaque câble
        Do While Not MaTableRow Is Nothing
            lKabelRef = MaTableRow.Value(pTTable5.FindField("KABEL_REF"))
            
            Dim pTLine4 As ITable
            Set pTLine4 = pFCLine4
            Dim pRRow As IRow
            Set pRRow = pTLine4.GetRow(lKabelRef)
            lKabelSpannung = pRRow.Value(pTLine4.FindField("SPANNUNG"))
            
            'Création du point de câble dans la coupe
            Dim pPointK As IPoint
            Dim lDistK As Double
            Dim lAngleK As Double
            
            lDistK = pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000 * lFactor - 0.07
            If lDistK < 0 Then
                lDistK = 0
            End If
            'MsgBox lDistK
            lAngleK = 2 * Pi / lKabelTotal
            Set pPointK = New Point
            
            pPointK.X = pPoint.X - lDistK * Sin(-lAngleC + lKabelCount * lAngleK + Pi)
            pPointK.Y = pPoint.Y - lDistK * Cos(-lAngleC + lKabelCount * lAngleK + Pi)
            
            Set pNewFeature6 = pFCPoint6.CreateFeature
            Set pNewFeature6.Shape = pPointK
            
            Set pFCursor6 = pFCPoint6.Search(Nothing, False)
            pNewFeature6.Value(pFCursor6.FindField("SPANNUNG")) = lKabelSpannung
            pNewFeature6.Value(pFCursor6.FindField("KABEL_REF")) = lKabelRef
            'pFCursor6.UpdateFeature pNewFeature6
            pNewFeature6.Store
            
            'MsgBox lRohrRef & "/" & lKabelRef & "/" & lKabelSpannung
            
            lKabelCount = lKabelCount + 1
            Set MaTableRow = MaTableCursor.NextRow
        Loop 'Fin parcours câble
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Création tubes internes
        Do While Not pFeature4 Is Nothing
            
            'Création du point de câble dans la coupe
            lDistK = pFeature2.Value(pFeature2.Class.FindField("DIAMETRE_MM")) / 2000 * lFactor - 0.07
            If lDistK < 0 Then
                lDistK = 0
            End If
            lAngleK = 2 * Pi / lKabelTotal
            Set pPointK = New Point
            
            pPointK.X = pPoint.X - lDistK * Sin(-lAngleC + lKabelCount * lAngleK + Pi)
            pPointK.Y = pPoint.Y - lDistK * Cos(-lAngleC + lKabelCount * lAngleK + Pi)
            
            Set pFeature4.Shape = pPointK
            pFeature4.Store
            
            lKabelCount = lKabelCount + 1
            Set pFeature4 = pFCursor4.NextFeature
        Loop 'Fin parcours câble
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        'MsgBox pPoint.X & "/" & pPoint.Y
        
'        Set pNewFeature = pFCPoint3.CreateFeature
'        Set pNewFeature.Shape = pPoint
'        pNewFeature.Store

        Set pFeature2.Shape = pPoint
        
        'Ajout du facteur correct pour affichage dans les attributs du tube
        pFeature2.Value(pFeature2.Class.FindField("FACTEUR")) = 2 / lFactor
        
        pFeature2.Store
        
        Set pFeature2 = pFCursor2.NextFeature
    Loop 'Fin parcours tube
        
    Set pFeature1 = pFCursor1.NextFeature
  Loop 'Fin parcours coupe

  'Stop feature editing
  pEditor.StopOperation ("Add Points")
  
  'Clear all feature selections
  pMap.ClearSelection
  
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

Public Function TrouveTable(nomtable As String) As ITable

    Dim pMxDoc As IMxDocument ' Un document
    Dim pMap As IMap ' Une carte
    Set pMxDoc = ThisDocument ' Le document est celui ouvert
    Set pMap = pMxDoc.FocusMap ' La carte est la carte active
    
    Dim pStTabColl As IStandaloneTableCollection ' Une collection de tables "libres"
    Dim pStTab As IStandaloneTable ' Une table libre
    Dim intcount As Long ' Un compteur
    Set pStTabColl = pMap ' La collection de table est celle de la carte active
    
    ' Boucle sur la collection de tables de la fenêtre active
    For intcount = 0 To pStTabColl.StandaloneTableCount - 1
    
        ' Si on trouve la table "NomTable on la stocke dans TrouveTable
        If (pStTabColl.StandaloneTable(intcount).Name = nomtable) Then
            Set pStTab = pStTabColl.StandaloneTable(intcount)
            Set TrouveTable = pStTab.Table
            Exit For
        End If
    Next
    ' Si on ne trouve pas la table erreur et on sort
    If TrouveTable Is Nothing Then
        MsgBox "Table " & nomtable & " introuvable", vbExclamation, "Erreur!"
        Exit Function
    End If
    
End Function ' Fin TrouveTable


Public Function AttributeQuery(pTable As esriGeoDatabase.ITable, Optional whereClause As String = "") As esriGeoDatabase.ICursor

  Dim pQueryFilter As esriGeoDatabase.IQueryFilter
  Dim pCursor As esriGeoDatabase.ICursor
  ' create a query filter
  Set pQueryFilter = New esriGeoDatabase.QueryFilter

  ' create the where statement
  pQueryFilter.whereClause = whereClause

  ' query the table passed into the function and use a cursor to hold the results
  Set pCursor = pTable.Search(pQueryFilter, False)

  Set AttributeQuery = pCursor
 
End Function





