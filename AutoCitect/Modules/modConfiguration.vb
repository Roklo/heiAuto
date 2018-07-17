Option Explicit On

Module modConfiguration

    Public TN, DES, AG, BG, ALD, ALL, AL, AH, AHH, NODE, SN, IOCNO, CH, RMin, RMax As String
    Public Emin, EMax, IasTN, Adr, Adr2, MbAdd, AYN, UN, ST, SANS, SANP, Aker_PATH, TNO, ISA, SYSN, NVAL, MaxAlm As String

    Public ColumnAlloc_Label, ColumnAlloc_ToolTip As String
    Public IOList_Path As String
    Public IOLIST_Sheet As String

    Private m_Dialog As frmConfigure


    Public Function ColumnAlloc_SFI_NUMBER() As String
        ColumnAlloc_SFI_NUMBER = TN
    End Function


    Public ReadOnly Property ColumnAlloc_TagNo() As String
        Get
            ColumnAlloc_TagNo = TNO
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_Description() As String
        Get
            ColumnAlloc_Description = DES
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_MODBUS_ADDRESS() As String
        Get
            ColumnAlloc_MODBUS_ADDRESS = MbAdd
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_NODE() As String
        Get
            ColumnAlloc_NODE = NODE
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_IASTagName() As String
        Get
            ColumnAlloc_IASTagName = IasTN
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_EngMin() As String
        Get
            ColumnAlloc_EngMin = Emin
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_EngMax() As String
        Get
            ColumnAlloc_EngMax = EMax
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_Unit() As String
        Get
            ColumnAlloc_Unit = UN
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_System() As String
        Get
            ColumnAlloc_System = SYSN
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_AlarmYesNo() As String
        Get
            ColumnAlloc_AlarmYesNo = AYN
        End Get
    End Property

    Public ReadOnly Property ColumnAlloc_SignalType() As String
        Get
            ColumnAlloc_SignalType = modConfiguration.ST
        End Get
    End Property

    Public Sub LoadConfiguration()
        If m_Dialog Is Nothing Then m_Dialog = New frmConfigure
        m_Dialog.LoadConfiguration
    End Sub

    Public Sub showConfigDialog()
        m_Dialog = New frmConfigure
        m_Dialog.ShowDialog
    End Sub

End Module
