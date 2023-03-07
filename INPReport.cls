VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CINPReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'...Container...
Private myContainer As ExtnUnitOperationContainer
Private m_InpExtnUtils As ExtnUtils_v2.CExtnUtils
'...DynContainer...
Private dyn_Container As ExtnDynUnitOpContainer

'****************************************************************************
'
'                                 EDF Variables
'
'****************************************************************************
Private edfFullStreamNameList As InternalTextFlexVariable
Private edfMBAlist As InternalRealFlexVariable

Private edfInStreamName As InternalTextVariable
Private edfInStreamNameList As InternalTextFlexVariable
Private edfInStreamFlowrateList As InternalRealFlexVariable
Private edfOutStreamName As InternalTextVariable
Private edfOutStreamNameList As InternalTextFlexVariable
Private edfOutStreamFlowrateList As InternalRealFlexVariable

Private edfLostFlowrateList As InternalRealFlexVariable
Private edfLostPercentList As InternalRealFlexVariable
Private edfLostAggregateList As InternalRealFlexVariable
Private edfLostAggregate As InternalRealVariable
Private edfTritiumInventoryList As InternalRealFlexVariable
Private edfTritiumInventoryTotal As InternalRealVariable
Private edfNetTritiumFlowInList As InternalRealFlexVariable

Private edfCompLabelName As InternalTextFlexVariable
Private edfStreamComp As InternalRealFlexVariable
Private edfComName As InternalTextFlexVariable

Private edfExtensionDLLVersion As InternalTextVariable
Private edfEngineDLLVersion As InternalTextVariable
Private edfUtilsDLLVersion As InternalTextVariable

'****************************************************************************
'
'                                  VB Variables
'
'****************************************************************************
Private Const M_H2 = 2.01588
Private Const M_HT = 4.02399
Private Const M_D2 = 4.028203556
Private Const M_DT = 5.030151056
Private Const M_T2 = 6.032098555
Private convTinHTgram As Double
Private convTinDTgram As Double
Private idxH2 As Long, idxHD As Long, idxHT As Long
Private idxD2 As Long, idxDT As Long, idxT2 As Long
Private nComp As Long

'... Global variables for MBA accounting ...
Private UserVarTAG As String
'   Declare next vectors as global to avoid resizing them each iteration in dynamics
Private inflowrate() As Double, outflowrate() As Double, permflowrate() As Double
Private lostflowrate() As Double, netinflowrate() As Double
Private lostpercentage() As Double, aggregatelost() As Double, aggregatelosttotal As Double
Private tritiuminventorytotal As Double, tritiuminventory() As Double, netinflowrateOld() As Double
Private streaminlist() As Variant, streamoutlist() As Variant, streampermlist() As Variant
Private OpNamesMBA() As String
Private OpActive() As Long      ' 0 == no user var, 1 == user var, -1 == no Op
Private flagGoal As Boolean     ' Indicates goal stream reached at some point in a MBA


Private nMBA As Long
'...Flags and other boolean variables...
Private IsForgetting As Boolean

Private opnames() As String
Private istream As ProcessStream
Private myfluid As Fluid
Private indexin As Long
Private indexout As Long
Private icompin As Long
Private Percent As Double
Private massTritium() As Double
Private TritiumPerc() As Double
Private TritiumPercName() As String
Private Complist() As String
Private CopyActive() As String

Private perdidasObject As Object
Private PerdidasFluid As Fluid

Private Flag As Boolean

Implements HYSYS.ExtnUnitOperation
Implements HYSYS.ExtensionObject


Private Function ExtensionObject_OnHelp(HelpPanel As String) As Boolean
    ExtensionObject_OnHelp = True
End Function
Private Sub ExtensionObject_Save()
End Sub
Private Sub ExtensionObject_StatusQuery(ByVal Status As HYSYS.ObjectStatus)
    ' Call m_InpExtnUtils.CheckInitType2("VRC")
    If myContainer.ExtensionInterface.IsIgnored = True Then Exit Sub
    If IsForgetting Then Exit Sub
End Sub
Private Sub ExtensionObject_VariableQuery(ByVal Variable As HYSYS.InternalVariableWrapper)
End Sub
Private Sub ExtnUnitOperation_BasisChanged()
End Sub
Private Sub ForgetResults()
End Sub


Private Function ExtensionObject_OnView(ViewName As String) As Boolean
    Call listStreams
    ' Call execute routine
    Call ExtnUnitOperation_Execute(False)
    Call accounting(UserVarTAG)
    ' Set states of tables to "Calculated" (black typing in HYSYS interface, not blue)
    Call setEDFmodifyState
    ExtensionObject_OnView = True
End Function

Private Function ExtnUnitOperation_Initialize(ByVal Container As HYSYS.ExtnUnitOperationContainer, ByVal IsRecalling As Boolean) As Long
 Dim IRV As InternalRealVariable '...for the controllers that only need to be initialized...
    '...Get the pointer to the container...
    Set myContainer = Container
    ' Get pointer to EDF variables
    Call PointEdfVariables
    ' Set states of tables to "Calculated" (black typing in HYSYS interface, not blue)
    Call setEDFmodifyState
    ' Get index for components of interest (based on basis manager) - permeating components
    Call iniCompIndex
    ' Initialize list of all streams in flowsheet
    Call listStreams
    ''' Get the pointer to Inprocess Extension Utils
    'Set m_InpExtnUtils = New CExtnUtils
    'Call m_InpExtnUtils.ExtnUtils_Initialize(Container, App)
    'm_InpExtnUtils.PutVersionInfo
    ' Define version number
    
    ' Define user variable tags used
    UserVarTAG = "PermeationT"
    convTinHTgram = (M_T2 / 2) / M_HT
    convTinDTgram = (M_T2 / 2) / M_DT
    
    If IsRecalling Then
        'm_InpExtnUtils.InitType = 2
        edfMBAlist.SetBounds UBound(edfMBAlist.State) + 1
        nMBA = UBound(edfMBAlist.State) + 1
    Else
        'm_InpExtnUtils.InitType = 1
        '...visibility controlers...
        Set IRV = myContainer.FindVariable("DesignEn").Variable
        IRV.value = 0
        Set IRV = myContainer.FindVariable("PerfomanceEn").Variable
        IRV.value = 0
        edfMBAlist.SetBounds 0
        edfInStreamNameList.SetBounds 0
        edfOutStreamNameList.SetBounds 0
    End If
    ExtnUnitOperation_Initialize = extnCurrentVersion
    Exit Function
ErrorCatch:
End Function

Private Sub ExtnUnitOperation_Execute(ByVal Forgetting As Boolean)
On Error GoTo ErrorCatch:
    '...Forget pass...
    If Forgetting Then
        IsForgetting = True
        ForgetResults
        Exit Sub
    End If
    Dim IRV As InternalVariableWrapper
    Set IRV = myContainer.FindVariable("DesignEn")
    '...Calulating pass...
    IsForgetting = False
    'FUTURO CHECKEO DE e.g QUE LOS STREAMS IN Y OUT CORRESPONDEN A LA MISMA LÍNEA
    If Not CheckInputs Then Exit Sub
    If Not GatherData Then
        Call myContainer.Trace(myContainer.Name & ": Error while Gathering Data in Execute.", False)
        Exit Sub
    End If
    If nMBA < 1 Then Exit Sub
    
    '*** MAIN SUB ***
    Call preAccounting
    '****************
    
    '...if we are here it is because we have solved the unit properly...
    myContainer.SolveComplete
    Exit Sub
ErrorCatch:
    MsgBox Err.Description
End Sub

Private Sub ExtensionObject_VariableChanged(ByVal Variable As HYSYS.InternalVariableWrapper)
On Error GoTo ERROR_CATCH
    Select Case Variable.tag
        Case "SetBounds"
            ' Make sure the new names are known
            If edfInStreamName.IsKnown And edfOutStreamName.IsKnown Then
                ' If we are here it means that we are increasing the number of MBA by one
                nMBA = nMBA + 1        '
                edfMBAlist.SetBounds nMBA
                edfMBAlist.SetValues genIndexArray(nMBA)
                ' Append new element to lists
                Call append(edfInStreamNameList, edfInStreamName.value)
                Call append(edfOutStreamNameList, edfOutStreamName.value)
                ' Pre-initialize vectors and EDF variables sizing once MBAs are known
                Call preAccounting
                ' Set states of tables to "Calculated" (black typing in HYSYS interface, not blue)
                Call setEDFmodifyState
            Else
                ' At least one of the new streams is not selected
                MsgBox ("MBA missing IN/OUT stream in operation " & myContainer.Name)
            End If
        Case "ClearMBA"
            ' Clear lists of in and out streams obtaining a fresh new T Report
            edfMBAlist.SetBounds 0
            nMBA = 0
            edfInStreamNameList.SetBounds 0
            edfOutStreamNameList.SetBounds 0
            edfInStreamFlowrateList.SetBounds 0
            edfOutStreamFlowrateList.SetBounds 0
            edfLostFlowrateList.SetBounds 0
            edfLostPercentList.SetBounds 0
            edfLostAggregateList.SetBounds 0
            edfTritiumInventoryList.SetBounds 0
            edfTritiumInventoryTotal.SetBounds 0
            edfNetTritiumFlowInList.SetBounds 0
        Case "ClearInventory"
            ' Set inventory variables and EDF-related variables to zero
            Dim i As Integer
            ReDim tritiuminventory(nMBA - 1)
            For i = 0 To nMBA - 1
                tritiuminventory(i) = 0
            Next
            tritiuminventorytotal = 0
            edfTritiumInventoryList.SetValues tritiuminventory
            edfTritiumInventoryTotal.SetValue tritiuminventorytotal
    End Select
    myContainer.SolveComplete
    Exit Sub
ERROR_CATCH:
    Call myContainer.Trace(myContainer.Name & ": Error in Variable Changed for variable " & Variable.tag & ".", False)
End Sub

Private Function ExtensionObject_VariableChanging(ByVal Variable As HYSYS.InternalVariableWrapper) As Boolean
Dim i As Long
    Select Case Variable.tag
    End Select
    ExtensionObject_VariableChanging = True
End Function

Private Sub ExtensionObject_Terminate()
'   m_InpExtnUtils.ExtnUtils_Terminate
   Set edfFullStreamNameList = Nothing
   Set edfMBAlist = Nothing
   
   Set edfInStreamName = Nothing
   Set edfInStreamNameList = Nothing
   Set edfInStreamFlowrateList = Nothing
   Set edfOutStreamName = Nothing
   Set edfOutStreamNameList = Nothing
   Set edfOutStreamFlowrateList = Nothing
   
   Set edfLostFlowrateList = Nothing
   Set edfLostPercentList = Nothing
   Set edfLostAggregateList = Nothing
   Set edfLostAggregate = Nothing
   Set edfTritiumInventoryList = Nothing
   Set edfTritiumInventoryTotal = Nothing
   Set edfNetTritiumFlowInList = Nothing
   
   Set edfCompLabelName = Nothing
   Set edfStreamComp = Nothing
   Set edfComName = Nothing
   
   Set edfExtensionDLLVersion = Nothing
   Set edfEngineDLLVersion = Nothing
   Set edfUtilsDLLVersion = Nothing
End Sub

Private Function CheckInputs() As Boolean
On Error GoTo ERROR_CATCH
    CheckInputs = True
    Exit Function
ERROR_CATCH:
    CheckInputs = False
End Function

Private Function GatherData() As Boolean
On Error GoTo Catch:
    GatherData = True
    Exit Function
Catch:
    GatherData = False
End Function


'***************************
'*** INITIALIZATION CODE ***
'***************************
Private Sub PointEdfVariables()
' Get objects of EDF variables
    With myContainer
        Set edfFullStreamNameList = .FindVariable("FullStreamNameList").Variable
        Set edfMBAlist = .FindVariable("MBAlist").Variable
        
        Set edfInStreamName = .FindVariable("InStreamName").Variable
        Set edfInStreamNameList = .FindVariable("InStreamNameList").Variable
        Set edfInStreamFlowrateList = .FindVariable("InStreamFlowrateList").Variable
        Set edfOutStreamName = .FindVariable("OutStreamName").Variable
        Set edfOutStreamNameList = .FindVariable("OutStreamNameList").Variable
        Set edfOutStreamFlowrateList = .FindVariable("OutStreamFlowrateList").Variable
        
        Set edfLostFlowrateList = .FindVariable("LostFlowrateList").Variable
        Set edfLostPercentList = .FindVariable("LostPercentList").Variable
        Set edfLostAggregateList = .FindVariable("LostAggregateList").Variable
        Set edfLostAggregate = .FindVariable("LostAggregate").Variable
        Set edfTritiumInventoryList = .FindVariable("TritiumInventoryList").Variable
        Set edfTritiumInventoryTotal = .FindVariable("TritiumInventoryTotal").Variable
        Set edfNetTritiumFlowInList = .FindVariable("NetTritiumFlowInList").Variable
        
        Set edfCompLabelName = .FindVariable("CompLabelName").Variable
        Set edfStreamComp = .FindVariable("StreamComp").Variable
        Set edfComName = .FindVariable("ComName").Variable
        
        Set edfExtensionDLLVersion = .FindVariable("ExtensionDLLVersion").Variable
        Set edfEngineDLLVersion = .FindVariable("EngineDLLVersion").Variable
        Set edfUtilsDLLVersion = .FindVariable("UtilsDLLVersion").Variable
    End With
End Sub
Private Sub setEDFmodifyState()
' Set modify states of EDF variables (mostly calculated, i.e. black letters)
    edfFullStreamNameList.SetModifyState vsCalculated
    edfMBAlist.SetModifyState vsCalculated
    
    edfInStreamName.SetModifyState vsCalculated
    edfInStreamNameList.SetModifyState vsCalculated
    edfInStreamFlowrateList.SetModifyState vsCalculated
    edfOutStreamName.SetModifyState vsCalculated
    edfOutStreamNameList.SetModifyState vsCalculated
    edfOutStreamFlowrateList.SetModifyState vsCalculated
    
    edfLostFlowrateList.SetModifyState vsCalculated
    edfLostPercentList.SetModifyState vsCalculated
    edfLostAggregateList.SetModifyState vsCalculated
    edfLostAggregate.SetModifyState vsCalculated
    edfTritiumInventoryList.SetModifyState vsCalculated
    edfTritiumInventoryTotal.SetModifyState vsCalculated
    edfNetTritiumFlowInList.SetModifyState vsCalculated

    edfCompLabelName.SetModifyState vsCalculated
    edfStreamComp.SetModifyState vsCalculated
    edfComName.SetModifyState vsCalculated
    
    edfExtensionDLLVersion.SetModifyState vsSpecified
    edfEngineDLLVersion.SetModifyState vsSpecified
    edfUtilsDLLVersion.SetModifyState vsSpecified
    edfExtensionDLLVersion.value = "6.0"
    'edfEngineDLLVersion = "1111"
    'edfUtilsDLLVersion = "2.0"
End Sub
Private Sub iniCompIndex()
' Get index for components that can permeate (based on basis manager)
  Dim ComponentList() As String
  Dim i As Long
    ComponentList = myContainer.Flowsheet.FluidPackage.Components.Names
    nComp = myContainer.Flowsheet.FluidPackage.Components.Count             ' Number of Components
    For i = 0 To nComp - 1
        Select Case ComponentList(i)
            Case "Hydrogen"
                idxH2 = i
            Case "HD*"
                idxHD = i
            Case "HT*"
                idxHT = i
            Case "Deuterium*"
                idxD2 = i
            Case "DT*"
                idxDT = i
            Case "Tritium*"
                idxT2 = i
        End Select
    Next i
End Sub
Private Sub listStreams()
' Saves a list of stream names in a string vector EDF variable
 Dim n As Long
 Dim stringList() As String
    n = myContainer.Flowsheet.MaterialStreams.Count
    ReDim stringList(n - 1)
    stringList = myContainer.Flowsheet.MaterialStreams.Names
    edfFullStreamNameList.SetBounds n
    edfFullStreamNameList.Values = stringList
End Sub
Private Sub iniVariables()
' Used to size vector global variables and EDF variables once the user has set the MBAs
 Dim i As Long
    '' Input check
    If nMBA < 1 Then Exit Sub
    '' Sizing/Bounding
    ReDim streaminlist(nMBA - 1), streamoutlist(nMBA - 1), streampermlist(nMBA - 1)
    ReDim inflowrate(nMBA - 1), outflowrate(nMBA - 1), permflowrate(nMBA - 1)
    ReDim lostflowrate(nMBA - 1), netinflowrate(nMBA - 1), netinflowrateOld(nMBA - 1)
    ReDim lostpercentage(nMBA - 1), aggregatelost(nMBA - 1)
    ReDim tritiuminventory(nMBA - 1)
    edfInStreamFlowrateList.SetBounds nMBA
    edfOutStreamFlowrateList.SetBounds nMBA
    edfLostFlowrateList.SetBounds nMBA
    edfLostPercentList.SetBounds nMBA
    edfLostAggregateList.SetBounds nMBA
    edfTritiumInventoryList.SetBounds nMBA
    edfNetTritiumFlowInList.SetBounds nMBA
    '' Get stream objects by passing its string name
    For i = 0 To nMBA - 1
        Set streaminlist(i) = myContainer.Flowsheet.MaterialStreams.Item(edfInStreamNameList(i))
        Set streamoutlist(i) = myContainer.Flowsheet.MaterialStreams.Item(edfOutStreamNameList(i))
    Next i
    '' Check empty values to initialize (This concerns Dynamic aggregation of losses)
    ' Get current list
    aggregatelost = edfLostAggregateList.Values
    tritiuminventory = edfTritiumInventoryList.Values
    netinflowrateOld = edfNetTritiumFlowInList.Values
    For i = 0 To nMBA - 1
        ' Check if some values are empty (-32767) and initialize them to zero
        If aggregatelost(i) = -32767 Then
            aggregatelost(i) = 0
        End If
        If tritiuminventory(i) = -32767 Then
            tritiuminventory(i) = 0
        End If
        If netinflowrateOld(i) = -32767 Then
            netinflowrateOld(i) = 0
        End If
    Next i
    ' Set current values of EDF to the updated ones
    edfLostAggregateList.SetValues aggregatelost
    edfTritiumInventoryList.SetValues tritiuminventory
End Sub




'***************************************
'*** CODE TO FORMAT THE INPUT MATRIX ***
'***************************************
Private Sub append(edfList As InternalTextFlexVariable, newElement As String)
' Appends one element to a FLEX EDF variable (probably doesn't work with attachment lists)
Dim n As Long   ' length
Dim newlist() As String
    ' Get old dimensions and old values.
    ' Note: If at the beginning the vector is empty, we jump over the next lines
    n = UBound(edfList.State) + 1
    If n > 0 Then
        ReDim newlist(n - 1)
        newlist = edfList.Values
    End If
    ' Expand dimensions by one and get new last element
    n = n + 1
    ReDim Preserve newlist(n - 1)
    newlist(n - 1) = newElement
    ' Set values to EDF variable
    edfList.SetBounds n
    edfList.Values = newlist
End Sub
Private Function genIndexArray(length As Long) As Long()
' Returns a vector of indices given its length: [1, 2, ..., length]
 Dim indexlist() As Long, i As Long
 ReDim indexlist(length - 1)
    For i = 0 To length - 1
        indexlist(i) = i + 1
    Next i
    genIndexArray = indexlist
End Function





'******************************
'*** TRITIUM pre-ACCOUNTING ***
'******************************
'   (Preparing steps for accounting)
Private Sub preAccounting()
' Call sequence of subroutines previous to accounting
    ' Initialize and size vectors
    Call iniVariables
    ' Loop over flowsheet to obtain a matrix of operations  -> matrix 'opNamesMBA'
    Call collectMBAdata
    ' Collect permeating streams for each MBA
    Call collectPermStreams
    ' Build operation-wise matrix from previous step        -> matrix 'OpActive'
    Call BuildActiveOpsMatrix(UserVarTAG)
End Sub
Private Sub collectMBAdata()
' DEEPER LOOP
' Loop over process lines (MASS BALANCE AREAS defined by IN stream and OUT stream in EDF)
' and assess tritium losses in them
 Dim currStream As Variant
 Dim currOp As Variant
 Dim i As Long, j As Long
On Error GoTo errorHandler
    '' Input check
    If nMBA < 1 Then Exit Sub
    ' Before looping, redim operations matrix
    ReDim OpNamesMBA(nMBA - 1, 0)
    For i = 0 To nMBA - 1
        flagGoal = False
        ' Initial stream for iterations
        Set currStream = myContainer.Flowsheet.MaterialStreams.Item(edfInStreamNameList(i))
        ' Call subroutine to start path
        Call Path(currStream, myContainer.Flowsheet.MaterialStreams.Item(edfOutStreamNameList(i)), i + 1)
        ' Check GOAL objective (endStream reached)
        If flagGoal = False Then
            MsgBox ("End stream could not be reached in Tritium Report (" + myContainer.Name + "). Check boundaries of MBA number " + CStr(i))
        End If
    Next i
    Exit Sub
errorHandler:
    MsgBox "Error in " + myContainer.Name + ". " + Err.Description
    'myContainer.SimulationCase.Solver.Integrator.Reset
End Sub
Private Sub collectPermStreams()
    ' Collect streams whose name starts by the number of the MBA: 10, 11, 12 and ends with 'out'
    Dim i As Long, j As Long, nStreams As Integer, stream As Variant, stageNum As String
    On Error GoTo errorHandler
    ' Input check
    If nMBA < 1 Then Exit Sub
    nStreams = myContainer.Flowsheet.MaterialStreams.Count
    For i = 0 To nMBA - 1
        ' Stringify stage num for search. 0 > 10, 1 > 11, 2 > 12
        stageNum = CStr(i + 10)
        For j = 0 To nStreams - 1
            Set stream = myContainer.Flowsheet.MaterialStreams.Item(j)
            If Right(stream.Name, 4) = "-out" And Left(stream.Name, 3) = stageNum + "-" Then
                Set streampermlist(i) = stream
            End If
        Next
    Next
    Exit Sub
errorHandler:
    MsgBox "Error in " + myContainer.Name + ". " + Err.Description
End Sub
Private Sub Path(currStream, endStream, MBA)
 Dim currOp As Variant, pendOp As Variant
 Dim j As Long, pathNum As Long, caseFlag As Long
    ' The starting point is the stream in the 'streaminlist'
        ' Go over line starting with inlet stream with function linearPath
        caseFlag = linearPath(currStream, currOp, endStream, MBA)
        Select Case caseFlag
            Case 0
            ' Reached end of MBA
                '********REVISAR CUANDO DEJEMOS CAMINOS PENDIENTES
                ' por ejemplo: If Not caminos_pendientes Then Exit For
                flagGoal = True
            Case 1
            ' Boundary reached, but not the goal end stream
                ' SI NO HEMOS ALCANZADO EL END STREAM EN OTRO CAMINO, HARÍA FALTA UNA FLAG
                ' DE AVISO AL USUARIO
            Case 2
            ' REACHED THE SAME BRANCH AS BEFORE. IT IS OK TO EXIT THIS PATH
            ' ANALIZAR SI VARIOS CASE TIENEN LA MISMA ACCIÓN COMO RESULTADO Y JUNTARLOS (1 Y 2)
            Case 3
            ' More than 2 streams in next operation
                ' First we account the number of paths and save the last operation
                Set pendOp = currOp
                pathNum = pendOp.AttachedProducts.Count
                ' Loop over the products of the last operation
                ' and call the 'Path' function again in order to analize each brach
                For j = 0 To pathNum - 1
                    If pendOp.AttachedProducts.Item(j).Name = endStream.Name Then
                        ' Raise flag and avoid continuing the path where the objective stream has already been discovered
                        flagGoal = True
                    Else
                        Call Path(pendOp.AttachedProducts.Item(j), endStream, MBA)
                    End If
                Next j
            Case 4
        End Select
End Sub
Private Function linearPath(ByRef currStream, ByRef currOp, ByVal endStream, _
                            ByVal currMBA As Long) As Long
' Go over a process linear line given the starting stream 'currStream'
' and current MBA 'currMBA' for accounting purposes
' Then, we store all operations inside the path in the matrix 'OpNamesPath'
' If 'endStream' is reached, we also exit this routine
'
' Returns INTEGER Flag:
' caseFlag = 0 -> Reached the objective 'endStream'
' caseFlag = 1 -> A boundary has been reached (feederblock)
' caseFlag = 2 -> We have reached a previously stored operation (loop). Later we can discard the rest of the line
' caseFlag = 3 -> The current operation has more than one output material streams
'
' being currStream and currOp defined as "ByRef", outside the function one can access the latest objects accessed
'
 Dim k As Long, L As Long, caseFlag As Long, nstream As Long
    For k = 1 To 1000#
        ' Get next operation
        Set currOp = currStream.DownstreamOpers.Item(0)
        '' CASE 0 (straight to the end)
        caseFlag = 0
        '' CASE 1
        ' Check if we came to the end without finding the expected OUT stream
        If currOp.TypeName = "feederblock" Then
            caseFlag = 1
            Exit For
        End If
        '' CASE 2
        ' Check if we have previously gone across this operation in the matrix operation,
        ' otherwise we add a new operation to the list with the function 'IsAlready'
        If IsAlready(currOp.Name, OpNamesMBA, True, currMBA) Then
            caseFlag = 2
            Exit For
        End If
        '' CASE 3
        ' Check if we have more than one output stream, so the line is not linear anymore
        If currOp.AttachedProducts.Count > 1 Then
            nstream = 0
            ' First discard energy streams
            For L = 0 To currOp.AttachedProducts.Count - 1
                If Not currOp.AttachedProducts.Item(L).IsEnergyStream Then
                    nstream = nstream + 1
                End If
            Next L
            If nstream > 1 Then
                ' The linear path ends here
                caseFlag = 3
                Exit For
            End If
        End If
        ''*********** FUTURE CHECKS
        '' CHECK USER VARIABLE
        '    If checkUserVar(currOp, UserVarTAG) Then
        '        lostflowrate(i) = lostflowrate(i) + currOp.GetUserVariable(UserVarTAG).Variable.value
        '    End If
        '' GET NEXT STREAM
        Set currStream = currOp.AttachedProducts.Item(0)
        ' Check if we have reached the end of the Area
        If currStream.Name = endStream.Name Then Exit For 'flagGoal will be raised later
    Next k
    linearPath = caseFlag
End Function
Private Sub BuildActiveOpsMatrix(tag As String)
' Loops over operations matrix and creates an integer matrix labelling operations
'   0 == No User Variable of tag 'tag'
'   1 == Existing User Variable of tag 'tag'
'   -1 == no Operation (empty space)
 Dim n As Long, m As Long, i As Long, j As Long
 Dim currOp As Variant
    '' Input check
    If nMBA < 1 Then Exit Sub
    '' Get the limits of the matrix
    n = UBound(OpNamesMBA, 1) + 1
    m = UBound(OpNamesMBA, 2) + 1
    ReDim OpActive(n - 1, m - 1)
    ' Redimension objective boolean matrix (considered as a global variable)
    For i = 0 To n - 1
        For j = 0 To m - 1
            ' Check if there is an empty space in the operation matrix
            If OpNamesMBA(i, j) = "" Then
                OpActive(i, j) = -1
            Else
                Set currOp = myContainer.SimulationCase.Flowsheet.Operations.Item(OpNamesMBA(i, j))
                If checkUserVar(currOp, tag) Then
                    OpActive(i, j) = 1
                Else
                    OpActive(i, j) = 0
                End If
            End If
        Next j
    Next i
End Sub




'**************************
'*** TRITIUM ACCOUNTING ***
'**************************
Private Sub accounting(myTag As String)
' Global accounting
 Dim currOp As Variant, currUserVar As Variant
 Dim i As Long, j As Long, n As Long, m As Long
 
On Error GoTo errorHandler
    '' Input check
    If nMBA < 1 Then Exit Sub
    
    '' Get limits of the matrix
    n = UBound(OpNamesMBA, 1) + 1
    m = UBound(OpNamesMBA, 2) + 1
    ' Brief size check
    If n <> nMBA Then
        MsgBox "sizes of MBA and matrix of operations do not match"
        Exit Sub
    End If
    
    ' GENERAL LOOP (overall values)
    ' Loop through each MBA and get their streams objects
    For i = 0 To nMBA - 1
        ' First account user variables
        lostflowrate(i) = 0
        ' CURRENTLY NOT ACCOUNTING FOR UNDESIRED LOSSES ******
        'For j = 0 To m - 1
        '    ' Accumulate values of permeated tritium user variables
        '    If OpActive(i, j) = 1 Then
        '        Set currOp = myContainer.SimulationCase.Flowsheet.Operations.Item(OpNamesMBA(i, j))
        '        Set currUserVar = currOp.GetUserVariable(myTag).Variable
        '        ' Identify User Variable units and process it consequently (UnitConversionType is Enum type of integers)
        '        Select Case currUserVar.UnitConversionType
        '            Case 3      ' Molar Flow = kmol/s
        '                lostflowrate(i) = lostflowrate(i) + currUserVar.value * M_T2
        '            Case 4      ' Mass Flow = kg/s
        '                lostflowrate(i) = lostflowrate(i) + currUserVar.value
        '            Case Else   ' Otherwise
        '                MsgBox ("Some error occurred with user variable " + myTag + " units.")
        '                Exit Sub
        '        End Select
        '    End If
        'Next j
        ' *****************************************************
        ' Compute the remaining flow rates (mass, kg/s)
        '   use second parenthesis to avoid errors while trying to compound properties
        '   remember to do conversions to account only for tritium
        inflowrate(i) = streaminlist(i).ComponentMassFlowValue()(idxT2) _
                      + streaminlist(i).ComponentMassFlowValue()(idxHT) * convTinHTgram _
                      + streaminlist(i).ComponentMassFlowValue()(idxDT) * convTinDTgram
        outflowrate(i) = streamoutlist(i).ComponentMassFlowValue()(idxT2) _
                       + streamoutlist(i).ComponentMassFlowValue()(idxHT) * convTinHTgram _
                       + streamoutlist(i).ComponentMassFlowValue()(idxDT) * convTinDTgram
        permflowrate(i) = streampermlist(i).ComponentMassFlowValue()(idxT2) _
                        + streampermlist(i).ComponentMassFlowValue()(idxHT) * convTinHTgram _
                        + streampermlist(i).ComponentMassFlowValue()(idxDT) * convTinDTgram
        netinflowrate(i) = inflowrate(i) - outflowrate(i) - lostflowrate(i) - permflowrate(i)
        ' Definition of % Tritium lost = losses / inlet. [NOTE: product by 100 is needed for percentage]
        '   Avoid division by zero setting a negative number to alert (we do not want to stop all the simulation always)
        If inflowrate(i) = 0 Then
            lostpercentage(i) = -0
        Else
            lostpercentage(i) = lostflowrate(i) / inflowrate(i) * 100
        End If
    Next i
    
    ' Set EDF values. NOTE that internal flow units are MASS FLOW UNITS [kg/s]
    edfInStreamFlowrateList.SetValues inflowrate
    edfOutStreamFlowrateList.SetValues outflowrate
    edfLostFlowrateList.SetValues lostflowrate
    edfLostPercentList.SetValues lostpercentage
    Exit Sub
errorHandler:
    MsgBox "Error in " + myContainer.Name + ". " + Err.Description
End Sub
Private Sub DynamicAccounting(deltaT)
' Calculate values of aggregate MASS tritium inventory loss along simulation
' Needs 'timestep' to interpolate
Dim i As Long
    '' Input check
    If nMBA < 1 Then Exit Sub
    ' Get aggregate flow
    'aggregatelosttotal = 0
    tritiuminventorytotal = 0
    For i = 0 To nMBA - 1
        ' Dynamic computation of lost flows - CURRENTLY NOT ACCOUNTING FOR UNDESIRED LOSSES
        'aggregatelost(i) = aggregatelost(i) + lostflowrate(i) * deltaT
        'aggregatelosttotal = aggregatelosttotal + aggregatelost(i)
        ' Dynamic computation of tritium inventory
        '  Integrate the inventory using the trapezoidal rule
        tritiuminventory(i) = tritiuminventory(i) + deltaT * (netinflowrate(i) + netinflowrateOld(i)) / 2
        tritiuminventorytotal = tritiuminventorytotal + tritiuminventory(i)
    Next i
    If tritiuminventorytotal < 0 Then
        ' Since edfTritiumInventoryTotal is of type "Mass", it cannot be negative
        tritiuminventorytotal = 0
    End If
    netinflowrateOld = netinflowrate  ' store previous net flow for next time step
    ' Fix values to EDF variables
    'edfLostAggregateList.SetValues aggregatelost
    'edfLostAggregate.SetValue aggregatelosttotal
    edfTritiumInventoryList.SetValues tritiuminventory
    edfTritiumInventoryTotal.SetValue tritiuminventorytotal
    edfNetTritiumFlowInList.SetValues netinflowrate
End Sub





'***************************
'*** AUXILIARY FUNCTIONS ***
'***************************
Private Function max(ParamArray list()) As Variant
'Return the "highest" of list of values
'---(ParamArray es una lista de elementos separados por comas)
Dim i As Long
Dim currentMax As Variant
    currentMax = list(0)
    For i = 1 To UBound(list)
        If list(i) > currentMax Then
          currentMax = list(i)
        End If
    Next i
    max = currentMax
End Function
Private Function min(ParamArray list()) As Variant
'Return the "lowest" of list of values
'---(ParamArray es una lista de elementos separados por comas)
Dim i As Long
Dim currentMin As Variant
    currentMin = list(0)
    For i = 1 To UBound(list)
        If list(i) < currentMin Then
          currentMin = list(i)
        End If
    Next i
    min = currentMin
End Function
Private Function checkUserVar(myOp, myTag) As Boolean
' Returns true or false in case the operation 'myOp' has the
' user variable named 'myTag' or not, respectively
 Dim x As Variant
 On Error GoTo Handler
    Set x = myOp.GetUserVariable(myTag)
    checkUserVar = True
    Exit Function
Handler:
    checkUserVar = False
End Function
Private Function IsAlready(ByVal value As String, ByRef matrix, addValue As Boolean, Optional currMBA As Long) As Boolean
' Look for 'value' inside 'matrix' and add it if required by 'addValue' and this is not repeated
'   if the same name appears, returns   True
'   if not, returns                     False
 Dim n As Long, m As Long, i As Long, j As Long
 Dim bound() As Long
    '' Get the limits of the matrix
    n = UBound(matrix, 1) + 1
    m = UBound(matrix, 2) + 1
    ReDim bound(n - 1)
    For i = 0 To n - 1
        ' Initialize columns bound
        bound(i) = -1
        For j = 0 To m - 1
            If matrix(i, j) = value Then
                IsAlready = True
                Exit Function
            End If
            ' Check when the row finishes (not all the rows have the same length
            If matrix(i, j) = "" Then Exit For
            bound(i) = j    ' In the 'j' element the last value is stored
        Next j
    Next i
    IsAlready = False
    
    '' Add value to the matrix if demanded
    If Not addValue Then Exit Function ' exit if the user does not demand the extra work
    If bound(currMBA - 1) < m - 1 Then
        ' In case there are empty spaces in the matrix to fill
        matrix(currMBA - 1, bound(currMBA - 1) + 1) = value
    Else
        ' In case we need to expand the matrix
        ReDim Preserve matrix(n - 1, m) ' add 1 column
        matrix(currMBA - 1, m) = value
    End If
End Function



'Private Function comp(ByRef namestr() As String, ByRef icomp As Long, streamcomplist() As Double, Compname() As String) As Boolean
'Dim streamcomp() As Double
'Dim i As Long, j As Long
'Dim compstream() As Double
'Dim complabelist() As String
'' Labeling composition tab
'ReDim complabelist(0 To UBound(myContainer.Flowsheet.FluidPackage.Components.Names))
'For i = 0 To UBound(myContainer.Flowsheet.FluidPackage.Components.Names)
'    complabelist(i) = (myContainer.Flowsheet.FluidPackage.Components.Item(i))
'Next
'edfCompLabelName.SetBounds UBound(myContainer.Flowsheet.FluidPackage.Components.Names) + 1
'edfCompLabelName.Values = complabelist
'ReDim Preserve streamcomp(0 To UBound(myContainer.Flowsheet.FluidPackage.Components.Names))
'edfStreamComp.SetBounds UBound(myContainer.Flowsheet.FluidPackage.Components.Names) + 1, UBound(streamcomplist) + 1
'i = 0
'Do
''For i = 0 To (UBound(namestr))
'    Set istream = myContainer.Flowsheet.Streams.Item(namestr(i))
'    Compname(icomp) = istream.Name
'    If Not istream.IsEnergyStream Then
'        streamcomp = istream.ComponentMassFraction.Values
'        For j = 0 To UBound(streamcomp)
'            streamcomplist(icomp, j) = streamcomp(j)
'        Next
'        icomp = icomp + 1
'    End If
'    i = i + 1
'Loop Until i > (UBound(namestr))
'' Setting values
'edfStreamComp.Values = streamcomplist
'edfComName.Values = Compname
'comp = True
'End Function






'***********************
'*** DYNAMIC METHODS ***
'***********************

Public Sub DynInitialize(dContainer As ExtnDynUnitOpContainer, IsRecalling As Boolean, MyVersion As Long, HoldupExist As Boolean)
' Called to Initialize the extension.
    Set dyn_Container = dContainer
    Call preAccounting
    'HoldupExist = False
    MyVersion = extnCurrentVersion
End Sub

Public Function InitializeSystem(ForceInit As Boolean) As Boolean
' Called each time the integrator starts
    ' Do the pre-analysis of paths of MBAs
    InitializeSystem = True
End Function



Public Function NumberOfFlowEquations() As Long
' Called to get the number of Pressure Flow equations the extension contributes separately.
    NumberOfFlowEquations = 0
End Function
Public Function NumberOfPressBalEquations() As Long
' Called to get the number of Pressure Balance equations the extension contributes separately.
    NumberOfPressBalEquations = 0
End Function
Public Function NumberOfFlowBalEquations() As Long
' Called to get the number of Flow Balance equations the extension contributes separately.
    NumberOfFlowBalEquations = 0
End Function
Public Function NumberOfGeneralEquations() As Long
' Called to get the number of General equations the extension contributes separately.
    NumberOfGeneralEquations = 0
End Function



Public Function VariablesInGeneralEquations()
' Called to get the number of variables in different General equations extension contributes separately.
End Function
Public Function GeneralEquationVars(EquationIdx As Long)
' Called to get the variables in each General equation.
End Function
Public Sub PrepareToIterateOnGeneralEqns(Dtime As Double)
' Called at each step of integration right before pressure flow solver starts iteration to solve the set of equations.
End Sub
Public Sub UpdateGeneralEqDerivsAndRHS(Dtime As Double, Derivs, Rhs)
' Called at each iteration of the Pressure Flow Solver to update the derivatives and right hand side of General equations.
End Sub



Public Function PreProcessStates(Dtime As Double) As Boolean
' Called at each step of the integration just before Pressure Flow Solver starts to solve the set of equations.
    PreProcessStates = True
End Function
Public Function PostProcessStates(Dtime As Double) As Boolean
    ' Called at each step of the integration just after Pressure Flow Solver finishes solving the set of equations.
    ' Account tritium losses and balances
    ' (Low frequency as it is only informative and does not influence the simulation results)
    Call accounting(UserVarTAG)
    ' Call aggregate accounting
    Call DynamicAccounting(Dtime)
    ' Set states of tables to "Calculated" (black typing in HYSYS interface, not blue)
    PostProcessStates = True
End Function

Public Function StepEnergyExplicitly(Dtime As Double) As Boolean
' Called to do the energy integration on the extension.
    StepEnergyExplicitly = True
End Function

Public Function StepCompositionExplicitly(Dtime As Double) As Boolean
' Called to do the composition integration on the extension.
 On Error GoTo errorHandler
    Call setEDFmodifyState
    ' End function
    StepCompositionExplicitly = True
    Exit Function
    
errorHandler:
    ' First attempt to avoid infinite errors poping up. Stop integrator
    MsgBox ("Errror in operation " & myContainer.Name & ". " & Err.Description & ". Integrator will stop.")
    StepCompositionExplicitly = False
End Function
