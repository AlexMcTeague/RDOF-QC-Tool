Attribute VB_Name = "KMZ"
Function InitKML(ByVal title As String) As MSXML2.DOMDocument60
    ' Make a new XML document
    Dim kml As MSXML2.DOMDocument60
    Set kml = New MSXML2.DOMDocument60
    
    ' Load in some initial, boilerplate XML, as well as a starting folder
    kml.LoadXML ("<?xml version=""1.0"" encoding=""UTF-8""?>" & _
        "<kml xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'>" & _
        "<Folder><name>" & title & "</name><open>1</open>" & _
        "<Style id=""dot""><scale>1</scale><Icon><href>http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png</href></Icon><LabelStyle><scale>0.85</scale></LabelStyle></Style>" & _
        "</Folder>" & _
        "</kml>")
        
    ' Set document properties and return the document
    kml.SetProperty "SelectionNamespaces", "xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'"
    kml.SetProperty "SelectionLanguage", "XPath"
        
    Set InitKML = kml
End Function

Function MakeFolder(ByRef xmlDoc As MSXML2.DOMDocument60, ByVal Name As String) As IXMLDOMNode
    ' Create and return a new named folder node
    Dim newFolder As IXMLDOMNode
    Set newFolder = xmlDoc.createElement("Folder")
    newFolder.appendChild(xmlDoc.createElement("name")).Text = Name
    
    Set MakeFolder = newFolder
End Function

Function MakeDot(ByRef xmlDoc As MSXML2.DOMDocument60, ByVal Name As String, ByVal lat As String, ByVal lng As String, ByVal vis As String) As IXMLDOMNode
    ' Create and return a new node that uses the dot style that is defined in the KML initialization code
    Dim newDot As IXMLDOMNode
    Set newDot = xmlDoc.createElement("Placemark")
    newDot.appendChild(xmlDoc.createElement("styleUrl")).Text = "#dot"
    newDot.appendChild(xmlDoc.createElement("Point")).appendChild(xmlDoc.createElement("coordinates")).Text = lat & "," & lng & ",0"
    newDot.appendChild(xmlDoc.createElement("name")).Text = Name
    newDot.appendChild(xmlDoc.createElement("visibility")).Text = vis
    
    Set MakeDot = newDot
End Function

Sub AnalyzeKMZ()
    Dim Dash As Worksheet
    Set Dash = ThisWorkbook.Sheets("KMZ")
    Dim pathKMZ As String
    Dim uuidList As Dictionary
    Set uuidList = New Dictionary
    
    ' Get the path of the KMZ file
    pathKMZ = CStr(ThisWorkbook.Sheets("File Imports").[Path_KMZ_Report].Value)
    
    ' If there isn't a KMZ path available, end the macro and inform the user
    If pathKMZ = "" Then
        ThisWorkbook.Sheets("File Imports").Activate
        [Path_KMZ_Report].Select
        MsgBox ("path_KMZ_Report is not set. Please select a file, then try again.")
        End
    End If
    
    ' If the file doesn't exist at the selected path, end the macro and inform the user
    KMZName = dir(pathKMZ)
    If KMZName = "" Then
        ThisWorkbook.Sheets("File Imports").Activate
        [Path_KMZ_Report].Select
        MsgBox ("File doesn't exist at path_KMZ_Report. Please select a different file, then try again.")
        End
    End If
    
    ' Prompt the user for a filepath for the KML output
    OLTName = "OLTName"
    ' Name suggestion for properly named deliverables (uses OLT name)
    If KMZName Like "*_*_KMZ_FIBER*.kmz" Then
        temp = Split(KMZName, "_")
        OLTName = CStr(temp(1))
    ' Name suggestion for raw downloads from Magellan (uses Prism ID if the workspace name matches "CHR_PID_01" formatting)
    ElseIf KMZName Like "KMZ_*_*_*_*" Then
        temp = Split(KMZName, "_")
        OLTName = CStr(temp(2))
    End If
    suggestedName = OLTName & " QC KML.kml"
    ' Ask the user for the save location for the QC KML
    outputPath = Application.GetSaveAsFilename( _
        fileFilter:="KML Files (*.kml), *.kml", _
        title:="Choose KML output location", _
        InitialFileName:=suggestedName)
    ' If the user clicked Cancel, ask them whether they want to continue the macro
    If outputPath = False Then
        Response = MsgBox("You didn't select a save location for the QC KML;" & vbNewLine & "Would you like the KMZ macro to continue anyway?", vbYesNo)
        If Response = vbYes Then
            
        Else: End
        End If
    End If
    
    kmlTitleWithExt = Mid(outputPath, InStrRev(outputPath, "\") + 1)
    kmlTitle = Left(kmlTitleWithExt, Len(kmlTitleWithExt) - 4)

    ' Load the file from the given path, whether it's a KML or KMZ file
    Set inputKML = Load_From_KML_Or_KMZ(pathKMZ)

    ' Define which featureClasses we want to pull out, and what data we want out of each class (COORDINATES ARE AUTOMATICALLY INCLUDED)
    Dim dataByFeatureClass As Dictionary
    Set dataByFeatureClass = New Dictionary
    With dataByFeatureClass
        If Dash.CheckBoxes("Checkbox_KMZ_Hybrid").Value = 1 Then .add "hybrid", "uuid,name,model"
        If Dash.CheckBoxes("Checkbox_KMZ_Slack").Value = 1 Then .add "slackStorage", "uuid,length,model"
        If Dash.CheckBoxes("Checkbox_KMZ_Riser").Value = 1 Then .add "riser", "uuid,length"
        If Dash.CheckBoxes("Checkbox_KMZ_Fiber").Value = 1 Then .add "fiberCable", "uuid,name,length,model"
        If Dash.CheckBoxes("Checkbox_KMZ_SpliceCan").Value = 1 Then .add "spliceCan", "uuid,name,model"
        If Dash.CheckBoxes("Checkbox_KMZ_SupportCable").Value = 1 Then .add "supportCable", "uuid,length"
        If Dash.CheckBoxes("Checkbox_KMZ_Support").Value = 1 Then .add "support", "uuid,plantType"
    End With
    
    ' Clear existing KMZ tabs
    ' TODO: the tabs are already cleared later in the macro, fully wiping them isn't necessary
    ' Instead of calling Clear_KMZ, it would be better to delete only un-used KMZ tabs
    Call Clear_KMZ
    
    ' Make a dictionary that will be populated with the lists of XML nodes, separated by feature class
    Dim results As Dictionary
    Set results = New Dictionary
        
    ' Read in the nodes separated by feature class
    For Each fc In dataByFeatureClass.Keys
        results.add fc, inputKML.SelectNodes("//kml:Folder/kml:Placemark/kml:ExtendedData/kml:Data[@name='featureClass' and kml:value='" & fc & "']/../..")
    Next
    
    If Not kmlTitleWithExt = False Then
        ' Initialize a new KML to output as a part of the analysis
        Dim kml As MSXML2.DOMDocument60
        Set kml = InitKML(kmlTitle)
        
        ' Get the top level folder of the new KML
        Dim topLevel As IXMLDOMNode
        Set topLevel = kml.SelectSingleNode("//kml:Folder")
    End If
    
    
    ' Read each property of node in each feature class, make a new worksheet for it
    For Each fc In dataByFeatureClass.Keys
        ' If the sheet doesn't exist, make a new one. If it does, clear it
        wsName = "KMZ_" & fc
        If WorksheetExists(wsName) Then
            Set ws = ThisWorkbook.Sheets(wsName)
            ws.Cells.Clear
        Else
            Set ws = ThisWorkbook.Sheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = wsName
        End If
        
        Set topLeft = ws.Range("A1")
        props = Split(dataByFeatureClass(fc), ",")
        rowCounter = 0
        
        ' Set the LAT and LONG headers for each sheet ahead of time
        topLeft.Offset(0, UBound(props) + 1).Value = "COORDINATES"
        
        If Not kmlTitleWithExt = False Then
            ' Make the KML folder for this feature class
            Dim fcFolder As IXMLDOMNode
            Set fcFolder = MakeFolder(kml, fc)
            topLevel.appendChild fcFolder
        End If
        
        ' For each node, read each of the properties
        For Each NODE In results(fc)
            nodeUUID = NODE.SelectSingleNode("kml:ExtendedData/kml:Data[@name='uuid']/kml:value").Text
            If Not uuidList.Exists(nodeUUID) Then
                uuidList.add nodeUUID, True
                Dim lat As Double, lng As Double
                
                ' Apply the right coordinate-parsing function, depending on whether the node has a polygon or a line string
                If Not (NODE.SelectSingleNode("kml:Polygon") Is Nothing) Then
                    polygonCoords NODE.SelectSingleNode("kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates").Text, lat, lng
                ElseIf Not (NODE.SelectSingleNode("kml:LineString") Is Nothing) Then
                    LineStringCoords NODE.SelectSingleNode("kml:LineString/kml:coordinates").Text, lat, lng
                End If
                
                
                topLeft.Offset(rowCounter + 1, UBound(props) + 1).Value = Round(lat, 6) & ", " & Round(lng, 6)
                
                ' Use a counter to keep track of which column we are filling out
                colCounter = 0
                
                For Each prop In props
                    ' If this is the first row, create the column headers on this pass as well
                    If (rowCounter = 0) Then
                        topLeft.Offset(0, colCounter).Value = UCase(prop)
                        
                        If Not kmlTitleWithExt = False Then
                            ' Piggyback on this to make the folders for each property for this feature class in the KML
                            Dim propFolder As IXMLDOMNode
                            Set propFolder = MakeFolder(kml, UCase(prop))
                            fcFolder.appendChild propFolder
                        End If
                    End If
                    
                    ' Try to find the value for the current property, and put that in the sheet if it exists
                    Set v = NODE.SelectSingleNode("kml:ExtendedData/kml:Data[@name='" & prop & "']/kml:value")
                    If Not (v Is Nothing) Then
                        topLeft.Offset(rowCounter + 1, colCounter) = v.Text
                        
                        If Not kmlTitleWithExt = False Then
                            ' Make a new dot for each property
                            Dim dot As IXMLDOMNode
                            Set dot = MakeDot(kml, v.Text, lng, lat, "0")
                            
                            ' Find the folder with the right name for this property and add the dot to it
                            fcFolder.SelectSingleNode("Folder[name='" & UCase(prop) & "']").appendChild dot
                        End If
                    End If
                    
                    ' Go to the next column
                    colCounter = colCounter + 1
                Next
                
                ' Got down to the next row
                rowCounter = rowCounter + 1
            End If
        Next
        With ws.Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .AutoFilter
        End With
        ws.Columns.AutoFit
        ws.Columns("A").Hidden = True
        Application.DisplayAlerts = False
        If ws.Range("A1").Value = "" Then ws.Delete
        Application.DisplayAlerts = True
    Next
    
    ' Save the kml file to the given path
    If Not kmlTitleWithExt = False Then kml.Save (outputPath)
    
    Call CheckAllErrors_KMZ
End Sub

Sub CheckAllErrors_KMZ()
    Call Clear_ErrorDashboard("KMZ")
    ' Do error checks on tabs that begin with "KMZ_"
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Name
        Case "KMZ_hybrid"
            
        Case "KMZ_slackStorage"
            LR = ws.Range("D1").End(xlDown).row
            
        Case "KMZ_riser"
            
        Case "KMZ_fiberCable"
            
        Case "KMZ_spliceCan"
            
        Case "KMZ_supportCable"
            LR = ws.Range("C1").End(xlDown).row

            Set rng = ws.Range("C2:C" & LR)
            uniqueCoords = WorksheetFunction.Unique(rng)
            If UBound(uniqueCoords) < rng.Count Then
                For Each cell In rng
                    If WorksheetFunction.CountIfs(rng, cell.Value) > 1 Then
                        Call AddError("Error_KMZ_DuplicateCable", cell.Value, "Duplicated cable or unnecesaary Aerial vertexes", cell.Offset(0, 1))
                    End If
                Next cell
            End If
        Case "KMZ_support"
            LR = ws.Range("C1").End(xlDown).row
            
            Set rng = ws.Range("C2:C" & LR)
            uniqueCoords = WorksheetFunction.Unique(rng)
            If UBound(uniqueCoords) < rng.Count Then
                For Each cell In rng
                    If WorksheetFunction.CountIfs(rng, cell.Value) > 1 Then
                        Call AddError("Error_KMZ_DuplicateSupport", cell.Value, "Duplicated pole/PED", cell.Offset(0, 1))
                    End If
                Next cell
            End If
        End Select
    Next i
End Sub


Function LineStringCoords(ByRef raw As String, ByRef lat As Double, ByRef lng As Double)
    ' Split all of the vertices into an array (space-delimited)
    Dim lineCoords() As String
    lineCoords = Split(raw, " ")
    
    ' Get the highest index of the array of vertices
    vertexCount = UBound(lineCoords)
    
    ' If the highest vertex index is odd then we need to find the midpoint of the two middle-most vertices
    ' Otherwise, just use the middle vertex
    If (vertexCount Mod 2 = 1) Then
        Dim coord1() As String, coord2() As String
        
        ' Get the vertices on each side of the middle line
        coord1 = Split(lineCoords(WorksheetFunction.Floor(vertexCount / 2, 1)), ",")
        coord2 = Split(lineCoords(WorksheetFunction.Ceiling(vertexCount / 2, 1)), ",")
        
        ' Get the midpoint between the two coordinates
        lat = (CDbl(coord1(1)) + CDbl(coord2(1))) / 2
        lng = (CDbl(coord1(0)) + CDbl(coord2(0))) / 2
    Else
        Dim coord() As String
        
        ' Get the middle coordinate
        coord = Split(lineCoords(vertexCount / 2), ",")
        lat = coord(1)
        lng = coord(0)
    End If
End Function

Function polygonCoords(raw As String, ByRef lat As Double, ByRef lng As Double)
    ' Split all of the vertices into an array (space-delimited)
    Dim polyCoords() As String
    polyCoords = Split(raw, " ")

    ' Get coordinates from opposite corners of the polygon
    Dim coord1() As String, coord2() As String
    coord1 = Split(polyCoords(0), ",")
    coord2 = Split(polyCoords(2), ",")
    
    ' Average the coordinates on each axis
    lat = (CDbl(coord1(1)) + CDbl(coord2(1))) / 2
    lng = (CDbl(coord1(0)) + CDbl(coord2(0))) / 2
End Function

Sub ExtractKML(ByVal pathToKMZ As String, ByVal returnPath As String, Optional ByVal temp As String = "")
    Dim applicationObject As Object
    Set applicationObject = CreateObject("Shell.Application")
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    ' If a temporary working folder is not provided, create one in the same folder as the KMZ
    tempDirectory = ""
    If (temp = "") Then
        tempDirectory = fileSystemObject.CreateFolder(fileSystemObject.GetParentFolderName(pathToKMZ) & "\KML_CONVERSION_TEMP")
    Else
        tempDirectory = temp
    End If
    
    ' Copy the KMZ to the temp folder as a KMZ
    pathAsZip = tempDirectory & "\" & Replace(fileSystemObject.GetFileName(pathToKMZ), ".kmz", ".7z")
    fileSystemObject.CopyFile pathToKMZ, pathAsZip, True
    
    ' Loop through the items in the zip file and operate on the one that ends in ".kml"
    For Each f In applicationObject.Namespace(pathAsZip).Items
        If Right(f.path, 4) = ".kml" Then
            ' Copy the kml file to the temp folder with whatever name it had within the kmz
            applicationObject.Namespace(tempDirectory).CopyHere f.path, 20
            
            ' Copy the kml to the destination path, which includes the intended file name
            fileSystemObject.CopyFile tempDirectory & "\" & f.Name, returnPath, True
            
            ' Clean up the unnamed kml
            fileSystemObject.DeleteFile tempDirectory & "\" & f.Name
        End If
    Next
    
    ' Clean up the zip file
    fileSystemObject.DeleteFile pathAsZip
    
    ' If we made a temp folder, clean that up as well
    If (temp = "") Then
        fileSystemObject.DeleteFolder tempDirectory
    End If
End Sub

Sub Clear_KMZ()
    ' Delete all tabs that begin with "KMZ_"
    i = 1
    Do While i <= ThisWorkbook.Worksheets.Count
        Name = ThisWorkbook.Worksheets(i).Name
        If Name Like "KMZ_*" Then
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(Name).Delete
            Application.DisplayAlerts = True
        Else
            i = i + 1
        End If
    Loop
End Sub

Sub HideOrShow_KMZ_UUIDs()
    ' Hide or show the first column in all "KMZ_" tabs
    i = 1
    Do While i <= ThisWorkbook.Worksheets.Count
        Name = ThisWorkbook.Worksheets(i).Name
        If Name Like "KMZ_*" Then
            With ThisWorkbook.Worksheets(i).Columns("A")
                If .Hidden = False Then
                    .Hidden = True
                Else
                    .Hidden = False
                End If
            End With
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
End Sub

Function Load_From_KML_Or_KMZ(path As String) As MSXML2.DOMDocument60
    ' Takes a filepath argument, returns an XML object

    ' If a KML path was given, we can proceed without hassle!
    If path Like "*.kml" Then
        kmlPath = path
    ' If a KMZ path was given, the KML will need to be extracted
    ElseIf path Like "*.kmz" Then
        ' Establish the path to the temp folder
        tempPath = Left(path, InStrRev(path, "\"))
        tempFolderName = "QC_SHEET_TEMP"
        tempFullPath = tempPath & tempFolderName
    
        ' Create the temp folder
        CreateFolderPath (tempFullPath)
        
        ' Extract the KML into the temp folder
        kmlPath = tempFullPath & "\ripeKML.kml"
        ExtractKML path, kmlPath, tempFullPath
        
        ' Alert the user if the KML file is missing (this happens if the KMZ fails to unzip properly)
        If dir(kmlPath) = "" Then
            MsgBox "KMZ couldn't be extracted! To fix, install 7-Zip, or open the KMZ file in Google Earth and re-save it before using with the QC Tool."
            End
        End If
    Else
        MsgBox "Expected KMZ or KML file!" & vbNewLine & "Path: " & path
        End
    End If
    
    
    ' Declare variables
    Dim xmlObj As MSXML2.DOMDocument60
    Dim Namespace As String
    
    ' Set KML Standard Namespace
    ' Namespace = "xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2' xmlns:kml='http://www.opengis.net/kml/2.2' xmlns:atom='http://www.w3.org/2005/Atom'"
    Namespace = "xmlns:kml='http://www.opengis.net/kml/2.2'"
    
    ' Load XML
    Set xmlObj = New MSXML2.DOMDocument60
    Call xmlObj.SetProperty("SelectionNamespaces", Namespace)
    Call xmlObj.SetProperty("SelectionLanguage", "XPath")
    xmlObj.async = False: xmlObj.validateOnParse = False
    xmlObj.Load (kmlPath)
    
    ' Clean up the temp directory if one exists
    If Not IsEmpty(tempFullPath) Then
        Dim fileSystemObject As Object
        Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
        fileSystemObject.DeleteFolder tempFullPath
    End If

    ' Return the XML object
    Set Load_From_KML_Or_KMZ = xmlObj
End Function
