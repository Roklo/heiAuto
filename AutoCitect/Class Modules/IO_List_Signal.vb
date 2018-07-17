Option Explicit On
Imports Microsoft.Office.Interop


Public Class IO_List_Signal
    Public Enum SIGNAL_TYPE
        SIGNAL_TYPE_UNKNOWM
        SIGNAL_TYPE_DIGITAL
        SIGNAL_TYPE_ANALOG
    End Enum

    Private m_RowInExcelSheet As Long
    Private m_Sheet As Excel.Worksheet

    Public Description As String
    Public Alarm_Yes_No As String
    Public SignalType As String
    Public MODBUS_ADDRESS As String
    Public SFI_NUMBER As clsSFINumber
    Public NODE As String

    Public EngMin As String
    Public EngMax As String
    Public Unit As String

    Public IASTagname As String
    Public ToolTip As String
    Public Label As String

    Private m_SignalInUse As Boolean

    Public TagNo As String
    Public ISA As String


    Public SystemName As String



    Public AlarmGroup As String
    Public Block As String


    Public AlarmDelay As String
    Public AllimLL As String
    Public AllimL As String
    Public Allim_H As String
    Public Allim_HH As String

    Public Station_No As String
    Public IOCard_No As String
    Public Channel As String
    Public RAW_Min As String
    Public RAW_Max As String
    Public IASRangeMin As String
    Public IASRangeMax As String

    Public Adress As String
    'Public ModbusAddress  As String
    Public Address2 As String
    Public NormalValue As String






    Public Property SignalInUse() As Boolean
        Get
            SignalInUse = m_SignalInUse
        End Get
        Set(piVal As Boolean)
            m_SignalInUse = piVal
        End Set
    End Property


    Public Function GetPropertByName(piName As String) As String
        Select Case LCase(piName)
            Case "description"
                GetPropertByName = Description
            Case "alarm description"
                GetPropertByName = Description
            Case "equipment desc."
                GetPropertByName = Description
            Case "alarm_yes_no"
                GetPropertByName = Alarm_Yes_No
            Case "signaltype"
                GetPropertByName = SignalType
            Case "modbus_address"
                GetPropertByName = MODBUS_ADDRESS
            Case "sfi_number"
                GetPropertByName = SFI_NUMBER.SFI
            Case "tag name"
                GetPropertByName = SFI_NUMBER.SFI
            Case "tag_name"
                GetPropertByName = SFI_NUMBER.SFI
            Case "node"
                GetPropertByName = NODE
            Case "side"
                GetPropertByName = NODE
            Case "alarmno"
                GetPropertByName = IASTagname
            Case "engmin"
                GetPropertByName = EngMin
            Case "engmax"
                GetPropertByName = EngMax

        'mapping for "nonstandard" genie "bargraph_general"
            Case "maxvalue"
                GetPropertByName = EngMax
            Case "minvalue"
                GetPropertByName = EngMin
        '---------------------------------------------

            Case "unit"
                GetPropertByName = Unit
            Case "label"
                GetPropertByName = Label
            Case "tooltip"
                GetPropertByName = ToolTip
            Case "tag"
                GetPropertByName = IASTagname
            Case "alarm no"
                GetPropertByName = IASTagname
            Case "alarmno"
                GetPropertByName = IASTagname
            Case "alarmno:"
                GetPropertByName = IASTagname
            Case "alarm.no:"
                GetPropertByName = IASTagname
            Case "tag"
                GetPropertByName = IASTagname

        End Select


    End Function

    Public Function Read(piWS As Excel.Worksheet, ByRef cur_line As Long, piIoList As IO_List)
        m_RowInExcelSheet = cur_line
     m_Sheet = piWS
    
    Description = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_Description)

        Alarm_Yes_No = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_AlarmYesNo)

        MODBUS_ADDRESS = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_MODBUS_ADDRESS)

        SignalType = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_SignalType)

        SFI_NUMBER.SetSFI(piWS.Cells(cur_line, modConfiguration.ColumnAlloc_SFI_NUMBER))

        NODE = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_NODE)

        EngMin = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_EngMin)

        EngMax = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_EngMax)

        Unit = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_Unit)

        IASTagname = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_IASTagName)

        ToolTip = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_ToolTip)

        Label = piWS.Cells(cur_line, modConfiguration.ColumnAlloc_Label)

        TagNo = piWS.Cells(cur_line, modConfiguration.TN)

        AlarmGroup = piWS.Cells(cur_line, modConfiguration.AG)
        Block = piWS.Cells(cur_line, modConfiguration.BG)
        AlarmDelay = piWS.Cells(cur_line, modConfiguration.ALD)
        AllimLL = piWS.Cells(cur_line, modConfiguration.ALL)
        AllimL = piWS.Cells(cur_line, modConfiguration.AL)
        Allim_H = piWS.Cells(cur_line, modConfiguration.AH)
        Allim_HH = piWS.Cells(cur_line, modConfiguration.AHH)

        Station_No = piWS.Cells(cur_line, modConfiguration.SN)
        IOCard_No = piWS.Cells(cur_line, modConfiguration.IOCNO)
        Channel = piWS.Cells(cur_line, modConfiguration.CH)
        RAW_Min = piWS.Cells(cur_line, modConfiguration.RMin)
        RAW_Max = piWS.Cells(cur_line, modConfiguration.RMax)

        IASRangeMin = piWS.Cells(cur_line, modConfiguration.RMin)
        IASRangeMax = piWS.Cells(cur_line, modConfiguration.RMax)
        'ModbusAddress = piWS.Cells(cur_line, modConfiguration.MbAdd)
        Adress = piWS.Cells(cur_line, modConfiguration.Adr)
        Address2 = piWS.Cells(cur_line, modConfiguration.Adr2)
        NormalValue = piWS.Cells(cur_line, modConfiguration.NVAL)

    End Function

    Public Function SetColumnInExcelSheet(piColumn As String, piValue As String)
        If m_Sheet Is Nothing Then Exit Function
        m_Sheet.Cells(m_RowInExcelSheet, piColumn) = piValue
    End Function

    Public Function GetColumnInExcelSheet(piColumn As String) As String
        GetColumnInExcelSheet = m_Sheet.Cells(m_RowInExcelSheet, piColumn)
    End Function


    Public Function IsIdentical(piOtherSignal As IO_List_Signal, Optional check_ias As Boolean = True, Optional ByRef poDifferences As String = "") As Boolean
        If piOtherSignal Is Nothing Then Exit Function
        Dim lRetval As Boolean
        poDifferences = ""
        lRetval = True

        If Not UCase(Left(Description, Len(Description) - 1)) = UCase(Left(piOtherSignal.Description, Len(Description) - 1)) Then
            lRetval = False
            poDifferences = "Description"
        End If

        If Not Alarm_Yes_No = piOtherSignal.Alarm_Yes_No Then
            lRetval = False
            poDifferences = poDifferences + "," + "Alarm_Yes_No"
        End If
        If Not SignalType = piOtherSignal.SignalType Then
            lRetval = False
            poDifferences = poDifferences + "," + "SignalType"
        End If
        If check_ias Then
            If Not MODBUS_ADDRESS = piOtherSignal.MODBUS_ADDRESS Then
                lRetval = False
                poDifferences = poDifferences + "," + "MODBUS_ADDRESS"
            End If
        End If
        If Not SFI_NUMBER.SFI = piOtherSignal.SFI_NUMBER.SFI Then
            lRetval = False
            poDifferences = poDifferences + "," + "SFI"
        End If
        If Not NODE = piOtherSignal.NODE Then
            lRetval = False
            poDifferences = poDifferences + "," + "NODE"
        End If
        If Not EngMin = piOtherSignal.EngMin Then
            lRetval = False
            poDifferences = poDifferences + "," + "EngMin"
        End If
        If Not EngMax = piOtherSignal.EngMax Then
            lRetval = False
            poDifferences = poDifferences + "," + "EngMax"
        End If
        If Not Unit = piOtherSignal.Unit Then
            lRetval = False
            poDifferences = poDifferences + "," + "UNIT"
        End If
        If check_ias Then
            If Not IASTagname = piOtherSignal.IASTagname Then
                lRetval = False
                poDifferences = poDifferences + "," + "IASTagname"
            End If
        End If

        If check_ias Then
            If Not ToolTip = piOtherSignal.ToolTip Then
                lRetval = False
                poDifferences = poDifferences + "," + "ToolTip"
            End If
            If Not Label = piOtherSignal.Label Then
                lRetval = False
                poDifferences = poDifferences + "," + "Label"
            End If
        End If

        If Not TagNo = piOtherSignal.TagNo Then
            lRetval = False
            poDifferences = poDifferences + "," + "TagNo"
        End If
        If Not ISA = piOtherSignal.ISA Then
            lRetval = False
            poDifferences = poDifferences + "," + "ISA"
        End If
        If Not SystemName = piOtherSignal.SystemName Then
            lRetval = False
            poDifferences = poDifferences + "," + "SystemName"
        End If
        If Not AlarmGroup = piOtherSignal.AlarmGroup Then
            lRetval = False
            poDifferences = poDifferences + "," + "AlarmGroup"
        End If
        If Not Block = piOtherSignal.Block Then
            lRetval = False
            poDifferences = poDifferences + "," + "Block"
        End If
        If Not AlarmDelay = piOtherSignal.AlarmDelay Then
            lRetval = False
            poDifferences = poDifferences + "," + "AlarmDelay"
        End If
        If Not AllimLL = piOtherSignal.AllimLL Then
            lRetval = False
            poDifferences = poDifferences + "," + "AllimLL"
        End If
        If Not AllimL = piOtherSignal.AllimL Then
            lRetval = False
            poDifferences = poDifferences + "," + "AllimL"
        End If
        If Not Allim_H = piOtherSignal.Allim_H Then
            lRetval = False
            poDifferences = poDifferences + "," + "Allim_H"
        End If
        If Not Allim_HH = piOtherSignal.Allim_HH Then
            lRetval = False
            poDifferences = poDifferences + "," + "Allim_HH"
        End If
        If Not Station_No = piOtherSignal.Station_No Then
            lRetval = False
            poDifferences = poDifferences + "," + "Station_No"
        End If
        If Not IOCard_No = piOtherSignal.IOCard_No Then
            lRetval = False
            poDifferences = poDifferences + "," + "IOCard_No"
        End If
        If Not Channel = piOtherSignal.Channel Then
            lRetval = False
            poDifferences = poDifferences + "," + "Channel"
        End If

        If check_ias Then
            If Not RAW_Min = piOtherSignal.RAW_Min Then
                lRetval = False
                poDifferences = poDifferences + "," + "RAW_Min"
            End If
            If Not RAW_Max = piOtherSignal.RAW_Max Then
                lRetval = False
                poDifferences = poDifferences + "," + "RAW_Max"
            End If
            If Not IASRangeMin = piOtherSignal.IASRangeMin Then
                lRetval = False
                poDifferences = poDifferences + "," + "IASRangeMin"
            End If
            If Not IASRangeMax = piOtherSignal.IASRangeMax Then
                lRetval = False
                poDifferences = poDifferences + "," + "IASRangeMax"
            End If
            If Not Adress = piOtherSignal.Adress Then
                lRetval = False
                poDifferences = poDifferences + "," + "Adress"
            End If
            If Not MODBUS_ADDRESS = piOtherSignal.MODBUS_ADDRESS Then
                lRetval = False
                poDifferences = poDifferences + "," + "MODBUS_ADDRESS"
            End If
            If Not Address2 = piOtherSignal.Address2 Then
                lRetval = False
                poDifferences = poDifferences + "," + "Address2"
            End If
        End If


        If Not NormalValue = piOtherSignal.NormalValue Then
            lRetval = False
            poDifferences = poDifferences + "," + "NormalValue"
        End If


        IsIdentical = lRetval

    End Function



    Private Sub Class_Initialize()
        SFI_NUMBER = New clsSFINumber
    End Sub

    Public ReadOnly Property AlarmNo() As String
        Get
            Dim testmsg As Integer
            testmsg = MsgBox("Error: Public Property Get AlarmNo() As String : DO NOT USE THIS!!",
                                 vbOKOnly + vbCritical, "Error")
        End Get

    End Property

    Public ReadOnly Property SignalTypeEnum() As SIGNAL_TYPE
        Get
            SignalTypeEnum = clsReplaceRule.RuleContentsStatus.REPLACE_RULE_UNKNOWN
            If Len(SignalType) < 2 Then
                Exit Property
            End If

            If UCase(Left(SignalType, 2)) = "DI" Then
                SignalTypeEnum = SIGNAL_TYPE.SIGNAL_TYPE_DIGITAL
            End If

            If UCase(Left(SignalType, 2)) = "AI" Then
                SignalTypeEnum = SIGNAL_TYPE.SIGNAL_TYPE_ANALOG
            End If
        End Get
    End Property

    Public Function AlarmNumber() As String
        On Error GoTo ERR_HANDLER
        AlarmNumber = Mid(Me.IASTagname, 4, 4)
        Exit Function
ERR_HANDLER:
        AlarmNumber = "0"
    End Function

End Class
