Attribute VB_Name = "M00_Public_Module"
'********************************************************************************
'*
'*  ���� �a����
'*
'*-------------------------------------------------------------------------------
'*
'*  ���i���W���[��
'*
'********************************************************************************
'*
'*  Date    :   2017/07/25
'*
'*  Author  :   Hideki Kanamori
'*
'********************************************************************************
'*
'*  Remake
'*
'*  No. Date        Author                  Note
'*
'*-------------------------------------------------------------------------------
'*
'*  01  2018/12/14  Hideki Kanamori         �A�N�Z�X�̐��l�ۂߌ덷�Ή�
'*  02  2020/01/28  Hideki Kanamori         Windows64�r�b�g�Ή�(�������x�A�b�v)
'*  03  9999/99/99  **********************  ************
'*  04  9999/99/99  **********************  ************
'*  05  9999/99/99  **********************  ************
'*  06  9999/99/99  **********************  ************
'*  07  9999/99/99  **********************  ************
'*  08  9999/99/99  **********************  ************
'*  09  9999/99/99  **********************  ************
'*  10  9999/99/99  **********************  ************
'*
'********************************************************************************
Option Compare Database
Option Explicit

''2019/05/23 Add Start
Public Const Con_DBG_Mode = 1
''2019/05/24 Add Start


''2018/12/14 Add Start
Public Const Con_Num_G = 0.01
''2018/12/14 Add End

''2017/12/11 Add Start
Public Const Con_Proc_Wait_Msg = "�b�����҂��������i�������j" & vbCrLf & "�p�\�R��������A�s��Ȃ��ŉ������B" & vbCrLf & "Please Wait" & vbCrLf & "Don't Touch Me!"
''2017/12/11 Add End

''2017/12/01 Add Start
Public Const Con_Mouse_Wait = True
Public Const Con_Mouse_Nor = False
''2017/12/01 Add End

''2017/11/15 Add Start

''�J�����_�[�\���J�n��
''2019/07/22 Add Start
Public Const Con_Calender_Start_Day = 26
''2019/07/22 Add End
''2019/07/22 Delete Start
''Public Const Con_Calender_Start_Day = 21
''''2017/11/15 Add End
''2019/07/22 Delete End

''2017/12/14 Add Start
Public Const Con_Order_Point = 1
Public Const Con_Performance_Point = 2
Public Const Con_Delivery_Record_Point = 3
Public Const Con_Required_Amount_Point = 4
''2017/12/14 Add End

''2017/12/19 Add Start
Public Const Con_Proc_Mode_None = 0
Public Const Con_Proc_Mode_New = 10
Public Const Con_Proc_Mode_CopyNew = 15
Public Const Con_Proc_Mode_Update = 20
Public Const Con_Proc_Mode_Delete = 30
''2017/12/19 Add End

''2020/01/28 Add Start
#If VBA7 And Win64 Then
    ''Private Const Proc_Wait = 0.001
    Private Const Proc_Wait = 0.01
    ''Private Const Proc_Wait = 1
#Else
    Private Const Proc_Wait = 1
#End If
''2020/01/28 Add End

''2020/01/28 Delete Start
'2018/10/19 Add Start
''Private Const Proc_Wait = 1
''Private Const Proc_Wait = 0.5
'2018/10/19 Add End
''2020/01/28 Delete End

Public Ret As Variant
Public Errloop As Error

''2017/12/19 Add Start
Public Int_Mode As Integer
''2017/12/19 Add End

''2017/12/22 Add Start
Public Int_TubeMaterial_Mode As Integer
''2017/12/22 Add End

' Data Folder
Public Const sDataFolder = "C:\Program Files\Brother bPAC3 SDK\Templates\"

''2019/03/05 Test Start

''2020/02/18 Add Start
Dim Ret2 As Variant
''2020/02/18 Add End

Public Type SD
    X01 As String * 10
    X02 As String * 20
End Type


''2019/03/05 Test End



Public Function Fnc_SQL_Exec(Str_SQL As String, Optional Str_PG As String) As Integer
'********************************************************************************
'*
'*  SQL���s
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    Fnc_SQL_Exec = False

    On Error GoTo Err_Fnc_SQL_Exec

    DoEvents
    
    DoCmd.SetWarnings False
    
    DoEvents

    ''2019/05/23 Add Start
    Ret = Fnc_DebugPrintFile("Fnc_SQL_Exec:" & Str_SQL, Str_PG)
    ''2019/05/23 Add End

    DoCmd.RunSQL Str_SQL
    
    DoEvents
    
    DoCmd.SetWarnings True

    DoEvents

    Fnc_SQL_Exec = True

Exit_Fnc_SQL_Exec:

    On Error GoTo 0

    Exit Function

Err_Fnc_SQL_Exec:

    Select Case Err
        Case 3464
            MsgBox Err.Number & ":" & Err.Description
            Ret = Fnc_DebugPrintFile("Fnc_SQL_Exec:" & Err.Number & ":" & Err.Description, Str_PG)
            Resume Next
        Case 3156
            MsgBox Err.Number & ":" & Err.Description
            Ret = Fnc_DebugPrintFile("Fnc_SQL_Exec:" & Err.Number & ":" & Err.Description, Str_PG)
            Resume Next
        Case 3157
            MsgBox Err.Number & ":" & Err.Description
            Ret = Fnc_DebugPrintFile("Fnc_SQL_Exec:" & Err.Number & ":" & Err.Description, Str_PG)
            Resume Next
''        Case 3086
''            ''MsgBox Err.Number & ":" & Err.Description
''            Resume Next
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                    Ret = Fnc_DebugPrintFile("Fnc_SQL_Exec:" & Err.Number & ":" & Err.Description, Str_PG)
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Query_Exec(Str_Query_Name As String, Optional Str_PG As String = "UnKnown") As Integer
''Public Function Fnc_Query_Exec(Str_Query_Name As String, Str_PG As String) As Integer
'********************************************************************************
'*
'*  �N�G���[���s
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    Fnc_Query_Exec = False

    On Error GoTo Err_Fnc_Query_Exec

    DoEvents
    
    ''2019/05/23 Add Start
    Ret = Fnc_DebugPrintFile("Fnc_Query_Exec:" & Str_Query_Name, Str_PG)
    ''2019/05/23 Add End
    
    DoCmd.SetWarnings False
    
    DoEvents
    
    DoCmd.OpenQuery Str_Query_Name
    
    DoEvents
    
    DoCmd.SetWarnings True

    DoEvents

    Fnc_Query_Exec = True

Exit_Fnc_Query_Exec:

    On Error GoTo 0

    Exit Function

Err_Fnc_Query_Exec:

    Select Case Err
        ''3073
        Case 3073
            Resume Next
        Case 3086
            ''MsgBox Err.Number & ":" & Err.Description
            Resume Next
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            
            ''2019/05/23 Add Start
            Ret = Fnc_DebugPrintFile("Fnc_Query_Exec�i�ڍׁj:" & Err.Description, Str_PG)
            ''2019/05/23 Add End

            ''�P�b�҂�
            Ret = Fnc_Proc_Wait(1)

            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Product_No_Get() As String
'********************************************************************************
'*
'*  ���i�ԍ��̎擾����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    On Error GoTo Err_Fnc_Product_No_Get

    DoEvents
    
    Fnc_Product_No_Get = DFirst("Make_ProductNo", "QS02_TM01_03_Product_No_Make")

    DoEvents

Exit_Fnc_Product_No_Get:

    On Error GoTo 0

    Exit Function

Err_Fnc_Product_No_Get:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Get_Record_Count(ByVal Str_Table_Query_Name As String) As Long
'********************************************************************************
'*
'*  �e�[�u�����N�G�������m�F
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          Str_Table_Query_Name    :   ���s����e�[�u�����N�G����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    Dim Obj_RS          As Recordset

    Fnc_Get_Record_Count = -1

    On Error GoTo Err_Fnc_Get_Record_Count

    DoEvents
    
    Set Obj_RS = CurrentDb.OpenRecordset(Str_Table_Query_Name)
    
    DoEvents
    
    ''2019/04/09 Add Start
    If Obj_RS.EOF = False Then
        Obj_RS.MoveLast
        DoEvents
        Fnc_Get_Record_Count = Obj_RS.RecordCount
    Else
        Fnc_Get_Record_Count = 0
    End If
    ''2019/04/09 Add End

    ''2019/04/09 Delete Start
''    Obj_RS.MoveLast
''    DoEvents
''    Fnc_Get_Record_Count = Obj_RS.RecordCount
    ''2019/04/09 Delete End



Exit_Fnc_Get_Record_Count:

    Obj_RS.Close
    Set Obj_RS = Nothing

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Record_Count:

    Select Case Err
        ''�J�����g�E���R�[�h��(3021)
        Case 3021
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''            Resume Next
            Fnc_Get_Record_Count = 0
            GoTo Exit_Fnc_Get_Record_Count
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

''Public Function Fnc_Dir_File_Select(Optional Str_Def_Dir As String = CurrentProject.Path, Optional Str_Def_File_Name As String) As String
''
''    On Error Resume Next
''    '�ϐ���`
''    Dim intRet As Integer         '�_�C�A���O�p�ϐ�
''    Dim Str_GetFileName As String     '�t���p�X�̒l
''
''    Fnc_Dir_File_Select = ""
''
''    With Application.FileDialog(msoFileDialogOpen)
''        '�_�C�A���O�̃^�C�g����ݒ�
''        .Title = "�t�@�C�����J���_�C�A���O"
''        '�t�@�C���̎�ނ�ݒ�
''        .Filters.Clear
''        .Filters.Add "�b�r�u�t�@�C��", "*.csv"
''        .FilterIndex = 1
''        '�����t�@�C���I���������Ȃ�
''        .AllowMultiSelect = False
''        '�����p�X��ݒ�
''        .InitialFileName = Str_Def_Dir
''        '�_�C�A���O��\��
''        intRet = .Show
''
''        If intRet <> 0 Then
''          '�t�@�C�����I�����ꂽ�Ƃ�
''          '���̃t���p�X��Ԃ�l�ɐݒ�
''          Str_GetFileName = Trim(.SelectedItems.Item(1))
''        Else
''          '�t�@�C�����I������Ȃ���΃u�����N
''          Str_GetFileName = ""
''        End If
''    End With
''    '�I�����ꂽ�t���p�X���e�L�X�g�{�b�N�X�֕\��
''    Fnc_Dir_File_Select = Str_GetFileName
''
''End Function

''Public Function Fnc_FileSelect(Optional Str_Def_Folder As String = "")
''
''On Error GoTo ErrorHandler  '�G���[�������[�`�������s���܂��B
''
''    Dim Returnvalue As Variant
''    Dim strmsg As String
''    Returnvalue = SysCmd(acSysCmdAccessVer)
''    strmsg = "Access2002�A2003�łȂ����߁A���̋@�\�𗘗p�ł��܂���B"
''
''    'Access�̃o�[�W�����𒲂ׂ܂��B
''    'Access2000��10.0�AAccess2000��9.0,Access97��8.0,Access95��7.0��Ԃ��܂��B
''
''    DoEvents
''
''    Select Case Returnvalue
''        Case Is > "10.0"
''
''            Dim inttype As Integer
''            Dim varSelectedFile As Variant
''
''            '�t�@�C����I������ꍇ�́Amsofiledialogfilepicker
''            '�t�H���_�[��I������ꍇ�́Amsofiledialogfolderpicker
''            inttype = msofiledialogfilepicker
''
''            '�t�@�C���Q�Ɨp�̐ݒ�l���Z�b�g���܂��B
''            With Application.FileDialog(inttype)
''
''                '�_�C�A���O�^�C�g����
''                .Title = "�t�@�C���I���@By Microsoft Access Club"
''
''                '�t�@�C���̎�ނ��`���܂��B
''                .Filters.Clear
''                .Filters.Add "CSV �t�@�C��", "*.CSV"
''''                .Filters.Add "HTML �t�@�C��", "*.html"
''''                .Filters.Add "HTM�t�@�C��", "*.htm"
''                .Filters.Add "���ׂẴt�@�C��", "*.*"
''
''                '�����t�@�C���I�����\�ɂ���ꍇ��True�A�s�̏ꍇ��False�B
''                .AllowMultiSelect = False
''
''                '�ŏ��ɊJ���z���_�[�𓖃t�@�C���̃t�H���_�[�Ƃ��܂��B
''                If Str_Def_Folder = "" Then
''                    .InitialFileName = CurrentProject.Path
''                Else
''                    .InitialFileName = Str_Def_Folder
''                End If
''
''                If .Show = -1 Then '�t�@�C�����I�������΁@-1 ��Ԃ��܂��B
''                    For Each varSelectedFile In .SelectedItems
''                        Fnc_FileSelect = varSelectedFile
''                    Next
''                End If
''
''            End With
''
''    Case Else
''
''        MsgBox strmsg, vbOKOnly, "Microsoft Access Club"
''
''    End Select
''
''    DoEvents
''
''Exit Function
''
''ErrorHandler:
''
''    MsgBox "�\�����ʃG���[���������܂���" & Chr(13) & _
''            "�G���[�i���o�[�F" & Err.Number & Chr(13) & _
''            "�G���[���e�F" & Err.Description, vbOKOnly
''    End
''
''End Function

''Public Function Fnc_OfficeFileDialog(intCK As Integer, Optional Str_Def_Folder As String = "")
''
''On Error GoTo ErrorHandler  '�G���[�������[�`�������s���܂��B
''
''    Dim strmsg As String
''    strmsg = "Access2002�A2003�łȂ����߁A���̋@�\�𗘗p�ł��܂���B"
''
''    DoEvents
''
''    If Fnc_VersionCK = True Then
''
''        Dim FD As FileDialog '�I�u�W�F�N�g�֕ϐ������B
''        Dim inttype As Integer
''        Dim varSelectedFile As Variant
''        Dim strtitle As String
''        Dim CK As Boolean
''
''        '�t�H���_�[��I������ꍇ�́Amsofiledialogfolderpicker
''        '�t�@�C����I������ꍇ�́Amsofiledialogfilepicker
''        If intCK = 1 Then
''            inttype = msoFileDialogFolderPicker
''            strtitle = "�t�H���_�[�I��"
''            CK = False
''        Else
''            inttype = msofiledialogfilepicker
''            strtitle = "�t�@�C���I��"
''            CK = True
''        End If
''
''        '�t�@�C���Q�Ɨp�̐ݒ�l���Z�b�g���܂��B
''        Set FD = Application.FileDialog(inttype)
''
''        '�_�C�A���O�^�C�g����
''        FD.Title = strtitle & " Microsoft Access Club"
''
''        '�t�H���_�[�I�����͕����I����s��(False)�B�t�@�C���I�����͉\(True)�B
''        FD.AllowMultiSelect = CK
''
''        With FD
''            '�ŏ��ɊJ���z���_�[�𓖃t�@�C���̃t�H���_�[�Ƃ��܂��B
''            If Str_Def_Folder = "" Then
''                .InitialFileName = CurrentProject.Path
''            Else
''                .InitialFileName = Str_Def_Folder
''            End If
''            If intCK = 1 Then
''
''            Else
''                .Filters.Clear
''                .Filters.Add "CSV �t�@�C��", "*.CSV"
''            End If
''       End With
''
''        If FD.Show = -1 Then '�t�@�C�����I�������΁@-1 ��Ԃ��B
''            For Each varSelectedFile In FD.SelectedItems
''                Fnc_OfficeFileDialog = varSelectedFile
''            Next
''        End If
''
''        Set FD = Nothing '�ϐ����J�����܂��B
''
''    Else
''        MsgBox strmsg, vbOKOnly, "Microsoft Access Club"
''    End If
''
''    DoEvents
''
''Exit Function
''
''ErrorHandler:
''
''    MsgBox "�\�����ʃG���[���������܂���" & Chr(13) & _
''            "�G���[�i���o�[�F" & Err.Number & Chr(13) & _
''            "�G���[���e�F" & Err.Description, vbOKOnly
''    End
''
''End Function

'Access2002�A2003�ł���΁AFnc_VersionCK�v���V�[�W����True��Ԃ��܂��B
Private Function Fnc_VersionCK() As Boolean

    Dim Returnvalue As Variant

    'acSysCmdAccessVer�͒萔�ł��B
    
    DoEvents
    
    Returnvalue = SysCmd(acSysCmdAccessVer)
    
    DoEvents

    'Access�̃o�[�W�����𒲂ׂ܂��B
    'Access2003��11.0�AAccess2002��10.0�AAccess2000��9.0
    'Access97��8.0,Access95��7.0��Ԃ��܂��B

''    If Returnvalue = "10.0" Or Returnvalue = "11.0" Then
''        Fnc_VersionCK = True
''    Else
''        Fnc_VersionCK = False
''    End If

    Select Case Returnvalue
        Case Is > "10.0"
            Fnc_VersionCK = True
        Case Else
            Fnc_VersionCK = False
    End Select

End Function

Public Function Fnc_RoundDown(I_Data, S_Pnt) As Currency
'********************************************************************************
'*
'*  �w���l�؎̂āx����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          I_Data      �F  ���͒l
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True    :   ����I��
'*
'*          False   :   �X�V��
'*
'********************************************************************************
'

    Dim T As Currency
    Dim U As Currency

    Fnc_RoundDown = 0

    On Error GoTo Err_Fnc_RoundDown

    DoEvents
    
    If IsEmpty(I_Data) = True Or IsNull(I_Data) = True Or IsNumeric(I_Data) = False Then
        Fnc_RoundDown = 0
        Exit Function
    End If

    T = 10 ^ Abs(S_Pnt)
    If S_Pnt >= 0 Then
        U = Abs(I_Data) * T
        If Int(U) = U Then
            Fnc_RoundDown = I_Data
        Else
            Fnc_RoundDown = Sgn(I_Data) * Int(U) / T
        End If
    Else
        U = Abs(I_Data) / T
        Fnc_RoundDown = Sgn(I_Data) * Int(U) * T
    End If
    
    DoEvents

Exit_Fnc_RoundDown:

    On Error GoTo 0

    Exit Function


Err_Fnc_RoundDown:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_RoundUp(ByVal I_Data As Currency, Optional S_Pnt As Integer = 0) As Currency
'********************************************************************************
'*
'*  �w���l�؏グ�x����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          I_Data      �F  ���͒l
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True    :   ����I��
'*
'*          False   :   �X�V��
'*
'********************************************************************************
'
    Dim T As Currency
    Dim U As Currency

    Fnc_RoundUp = 0

    On Error GoTo Err_Fnc_RoundUp
    
    DoEvents
    
    If IsEmpty(I_Data) = True Or IsNull(I_Data) = True Or IsNumeric(I_Data) = False Then
        Fnc_RoundUp = 0
        Exit Function
    End If

    T = 10 ^ Abs(S_Pnt)
    If S_Pnt >= 0 Then
        U = Abs(I_Data) * T
        If Int(U) = U Then
            Fnc_RoundUp = I_Data
        Else
            Fnc_RoundUp = Sgn(I_Data) * Int(U + 1) / T
        End If
    Else
        U = Abs(I_Data) / T
        If Abs(I_Data) > Int(U) * T Then
            Fnc_RoundUp = Sgn(I_Data) * Int(U + 1) * T
        Else
            Fnc_RoundUp = Sgn(I_Data) * Int(U) * T
        End If
    End If

    DoEvents

Exit_Fnc_RoundUp:

    On Error GoTo 0

    Exit Function


Err_Fnc_RoundUp:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Round(ByVal I_Data As Currency, S_Pnt As Integer) As Currency
'********************************************************************************
'*
'*  �w�l�̌ܓ��x����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          I_Data      �F  ���͒l
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True    :   ����I��
'*
'*          False   :   �X�V��
'*
'********************************************************************************
'
    Dim T As Currency
    Dim U As Currency

    Fnc_Round = 0

    On Error GoTo Err_Fnc_Round

    DoEvents
    
    If IsEmpty(I_Data) = True Or IsNull(I_Data) = True Or IsNumeric(I_Data) = False Then
        Fnc_Round = 0
        Exit Function
    End If

    T = 10 ^ Abs(S_Pnt)
    If S_Pnt >= 0 Then
        U = Abs(I_Data) * T
        If U - Int(U) < 0.5 Then
            Fnc_Round = Sgn(I_Data) * Int(U) / T
        Else
            Fnc_Round = Sgn(I_Data) * Int(U + 1) / T
        End If
''    Else
''        U = Abs(I_Data) / T
''        If Abs(I_Data) > Int(U) * T Then
''            Fnc_Round = Sgn(I_Data) * Int(U + 1) * T
''        Else
''            Fnc_Round = Sgn(I_Data) * Int(U) * T
''        End If
    End If
    
    DoEvents

Exit_Fnc_Round:

    On Error GoTo 0

    Exit Function


Err_Fnc_Round:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function
Public Function Fnc_Get_Month_End(C_Year, C_Month) As Integer
'********************************************************************************
'*
'*  �����E�擾
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*           C_Year     :   �`�F�b�N�w�N�x
'*
'*           C_Month    :   �`�F�b�N�w���x
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'
    Dim NextMonth, EndOfMonth

    Fnc_Get_Month_End = False

    On Error GoTo Err_Fnc_Get_Month_End

    DoEvents
    
    NextMonth = DateAdd("m", 1, DateSerial(C_Year, C_Month, 1))

    EndOfMonth = NextMonth - DatePart("d", NextMonth)

    Fnc_Get_Month_End = DatePart("d", EndOfMonth)
    
    DoEvents

Exit_Fnc_Get_Month_End:

    On Error GoTo 0

    Exit Function


Err_Fnc_Get_Month_End:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Double_Start_Chk(Chk_DB) As Integer
'********************************************************************************
'*
'*   ��d�N���`�F�b�N
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*           Fld_Date
'*                   :   ���t�\���t�B�[���h
'*
'*           Fld_Time
'*                   :   ���ԕ\���t�B�[���h
'*
'********************************************************************************
'

    Fnc_Double_Start_Chk = False

    On Error GoTo Err_Fnc_Double_Start_Chk



    Fnc_Double_Start_Chk = True

Exit_Fnc_Double_Start_Chk:

    On Error GoTo 0

    Exit Function


Err_Fnc_Double_Start_Chk:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

''2017/11/07 Add Start
Public Function Fnc_Query_Open_Name(T_Name, DB_Open, DS_Open, OP_FLG) As Integer
'********************************************************************************
'*
'*  �N�G���[�E�I�[�v�������i�N�G���[���w��Ver�j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          T_Name      :   �e�[�u����
'*          DB_Open     :   �f�[�^�E�x�[�X��`
'*          DS_Open     :   ���R�[�h�E�Z�b�g��`
'*          OP_Flg      :   �t�@�C���E�I�[�v���E�t���O
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'

    Fnc_Query_Open_Name = False

    On Error GoTo Err_Fnc_Query_Open_Name

    DoEvents
    
    Set DB_Open = OpenDatabase(CurrentDb.Name)
    ''Set DS_Open = DB_Open.OpenRecordset(T_Name, dbOpenDynaset)             '�e�[�u���E�I�[�v��
    
    DoEvents
    
    Set DS_Open = DB_Open.OpenRecordset(T_Name)             '�e�[�u���E�I�[�v��

    OP_FLG = OP_FLG + 1

    Fnc_Query_Open_Name = True

Exit_Fnc_Query_Open_Name:

    On Error GoTo 0

    Exit Function


Err_Fnc_Query_Open_Name:

    Select Case Err
        '3078
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
            Resume
    End Select

End Function

Public Function Fnc_Query_Close(DB_Close, DS_Close, OP_FLG) As Integer
'********************************************************************************
'*
'*  �e�[�u���E�N���[�Y����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          DB_Close    :   �f�[�^�E�x�[�X��`
'*          DS_Close    :   ���R�[�h�E�Z�b�g��`
'*          OP_Flg      :   �t�@�C���E�I�[�v���E�t���O
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'

    Fnc_Query_Close = False

    On Error GoTo Err_Fnc_Query_Close

    DoEvents
    
    DS_Close.Close
    
    DoEvents
    
    DB_Close.Close
    
    DoEvents
    
    Set DS_Close = Nothing
    
    DoEvents
    
    Set DB_Close = Nothing

    OP_FLG = OP_FLG - 1

    Fnc_Query_Close = True

Exit_Fnc_Query_Close:

    On Error GoTo 0

    Exit Function


Err_Fnc_Query_Close:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function


Public Function Fnc_Query_Open_ADO(T_Name, DS_Open, OP_FLG) As Integer
'********************************************************************************
'*
'*  �N�G���[�E�I�[�v�������i�N�G���[���w��Ver�j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          T_Name      :   �e�[�u����
'*          DB_Open     :   �f�[�^�E�x�[�X��`
'*          DS_Open     :   ���R�[�h�E�Z�b�g��`
'*          OP_Flg      :   �t�@�C���E�I�[�v���E�t���O
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'

    Fnc_Query_Open_ADO = False

    On Error GoTo Err_Fnc_Query_Open_Name

    DoEvents
    
    DS_Open.Open T_Name, CurrentProject.Connection, , adLockOptimistic
    
    DoEvents
    
    OP_FLG = OP_FLG + 1

    Fnc_Query_Open_ADO = True
    
    DoEvents

Exit_Fnc_Query_Open_Name:

    On Error GoTo 0

    Exit Function


Err_Fnc_Query_Open_Name:

    Select Case Err
        '3078
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
            Resume
    End Select

End Function

Public Function Fnc_Query_Close_ADO(DS_Close, OP_FLG) As Integer
'********************************************************************************
'*
'*  �e�[�u���E�N���[�Y����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          DB_Close    :   �f�[�^�E�x�[�X��`
'*          DS_Close    :   ���R�[�h�E�Z�b�g��`
'*          OP_Flg      :   �t�@�C���E�I�[�v���E�t���O
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'

    Fnc_Query_Close_ADO = False

    On Error GoTo Err_Fnc_Query_Close

    DoEvents
    
    DS_Close.Close
    
    DoEvents
    
    Set DS_Close = Nothing
    
    DoEvents

    OP_FLG = OP_FLG - 1

    Fnc_Query_Close_ADO = True
    
    DoEvents

Exit_Fnc_Query_Close:

    On Error GoTo 0

    Exit Function


Err_Fnc_Query_Close:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

''2017/11/07 Add End

''2017/11/14 Add Start
Public Function Fnc_Get_This_Year() As Integer
    
    On Error GoTo Err_Fnc_Get_This_Year
    
    DoEvents
    
    Fnc_Get_This_Year = Year(Fnc_Get_This_YM())
    
    DoEvents

Exit_Fnc_Get_This_Year:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_This_Year:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_This_Month() As Integer
    
    On Error GoTo Err_Fnc_Get_This_Month

    DoEvents
    
    Fnc_Get_This_Month = Month(Fnc_Get_This_YM())

    DoEvents

Exit_Fnc_Get_This_Month:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_This_Month:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Next_Year_Month() As Date

    Dim Int_Year As Integer
    Dim Int_Month As Integer
    Dim Dte_Wk As Date

    On Error GoTo Err_Fnc_Get_Next_Year_Month
    DoEvents
    
    Int_Year = Fnc_Get_This_Year()
    Int_Month = Fnc_Get_This_Month()

    Dte_Wk = DateSerial(Int_Year, Int_Month + 2, 0)

''    Int_Next_Year = Year(Dte_Wk)
''    Int_Next_Month = Month(Dte_Wk)

    Fnc_Get_Next_Year_Month = Dte_Wk
    
    DoEvents

Exit_Fnc_Get_Next_Year_Month:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Next_Year_Month:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function
''2017/11/14 Add End

''2017/11/15 Add Start
Public Function Fnc_Get_This_YM() As Date

    Dim Int_Year As Integer
    Dim Int_Month As Integer
    Dim Int_Day As Integer
    Dim Dte_Wk As Date
    
    On Error GoTo Err_Fnc_Get_This_YM
    
    DoEvents

    ''���ݓ������Z�b�g
    Dte_Wk = Now()

    Int_Year = Year(Dte_Wk)
    Int_Month = Month(Dte_Wk)
    Int_Day = Day(Dte_Wk)

    Select Case Int_Day
        Case Is >= Con_Calender_Start_Day
            Int_Month = Int_Month + 1
        Case Else
    End Select

    Fnc_Get_This_YM = DateSerial(Int_Year, Int_Month, Int_Day)
    
    DoEvents

Exit_Fnc_Get_This_YM:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_This_YM:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_Next_Year_Month(ByVal Int_Year As Integer, ByVal Int_Month As Integer) As Date
''    Dim Int_Year As Integer
''    Dim Int_Month As Integer
    Dim Dte_Wk As Date
    
    On Error GoTo Err_Fnc_Get_Calc_Next_Year_Month
    
    DoEvents

    If Int_Year = 0 Then
        Int_Year = Fnc_Get_This_Year()
    End If

    If Int_Month = 0 Then
        Int_Month = Fnc_Get_This_Month()
    End If

'    Int_Year = Fnc_Get_This_Year()
'    Int_Month = Fnc_Get_This_Month()

    Dte_Wk = DateSerial(Int_Year, Int_Month + 2, 0)

''    Int_Next_Year = Year(Dte_Wk)
''    Int_Next_Month = Month(Dte_Wk)

    Fnc_Get_Calc_Next_Year_Month = Dte_Wk
    
    DoEvents

Exit_Fnc_Get_Calc_Next_Year_Month:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_Next_Year_Month:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function
''2017/11/15 Add End

''2017/11/21 Add Start
Public Function Fnc_Get_Calc_Next_Year(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Long
    
    On Error GoTo Err_Fnc_Get_Calc_Next_Year
    
    DoEvents
    
    Fnc_Get_Calc_Next_Year = Year(Fnc_Get_Calc_Next_Year_Month(Int_Year, Int_Month))

    DoEvents

Exit_Fnc_Get_Calc_Next_Year:
    
    On Error GoTo 0
    
    Exit Function

Err_Fnc_Get_Calc_Next_Year:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_Next_Month(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Long
    
    On Error GoTo Err_Fnc_Get_Calc_Next_Month
    
    DoEvents
    
    Fnc_Get_Calc_Next_Month = Month(Fnc_Get_Calc_Next_Year_Month(Int_Year, Int_Month))

    DoEvents

Exit_Fnc_Get_Calc_Next_Month:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_Next_Month:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_Next_Year_Fmt(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As String
    
    On Error GoTo Err_Fnc_Get_Calc_Next_Year_Fmt
    
    DoEvents
    
    Fnc_Get_Calc_Next_Year_Fmt = Format(Fnc_Get_Calc_Next_Year(Int_Year, Int_Month), "0000")

    DoEvents

Exit_Fnc_Get_Calc_Next_Year_Fmt:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_Next_Year_Fmt:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_Next_Month_Fmt(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As String

    On Error GoTo Err_Fnc_Get_Calc_Next_Month_Fmt

    DoEvents
    
    Fnc_Get_Calc_Next_Month_Fmt = Format(Fnc_Get_Calc_Next_Month(Int_Year, Int_Month), "00")

    DoEvents

Exit_Fnc_Get_Calc_Next_Month_Fmt:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_Next_Month_Fmt:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function
''2017/11/21 Add End
''2017/11/22 Add Start

Public Function Fnc_Get_Calc_This_YM(ByVal Int_Year As Integer, ByVal Int_Month As Integer) As Date
''    Dim Int_Year As Integer
''    Dim Int_Month As Integer
    Dim Dte_Wk As Date

    On Error GoTo Err_Fnc_Get_Calc_This_YM

    DoEvents
    
    If Int_Year = 0 Then
        Int_Year = Fnc_Get_This_Year()
    End If

    If Int_Month = 0 Then
        Int_Month = Fnc_Get_This_Month()
    End If

    Dte_Wk = DateSerial(Int_Year, Int_Month, 1)
    
    DoEvents

    Fnc_Get_Calc_This_YM = Dte_Wk

Exit_Fnc_Get_Calc_This_YM:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_This_YM:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_This_Year(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Integer
    
    On Error GoTo Err_Fnc_Get_Calc_This_Year
    
    DoEvents
    
    Fnc_Get_Calc_This_Year = Year(Fnc_Get_Calc_This_YM(Int_Year, Int_Month))

    DoEvents

Exit_Fnc_Get_Calc_This_Year:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_This_Year:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_This_Month(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Integer
    
    On Error GoTo Err_Fnc_Get_Calc_This_Month
    
    DoEvents
    
    Fnc_Get_Calc_This_Month = Month(Fnc_Get_Calc_This_YM(Int_Year, Int_Month))

    DoEvents

Exit_Fnc_Get_Calc_This_Month:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_This_Month:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_This_Year_Fmt(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Integer
    
    On Error GoTo Err_Fnc_Get_Calc_This_Year_Fmt
    
    DoEvents
    
    Fnc_Get_Calc_This_Year_Fmt = Fnc_Get_Calc_This_Year(Int_Year, Int_Month)

    DoEvents

Exit_Fnc_Get_Calc_This_Year_Fmt:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_This_Year_Fmt:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Get_Calc_This_Month_Fmt(Optional ByVal Int_Year As Integer = 0, Optional ByVal Int_Month As Integer = 0) As Integer
    
    On Error GoTo Err_Fnc_Get_Calc_This_Month_Fmt
    
    DoEvents
    
    Fnc_Get_Calc_This_Month_Fmt = Fnc_Get_Calc_This_Month(Int_Year, Int_Month)
    
    DoEvents

Exit_Fnc_Get_Calc_This_Month_Fmt:

    On Error GoTo 0
    
    Exit Function

Err_Fnc_Get_Calc_This_Month_Fmt:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

''Public Function Fnc_Get_Table_This_YM() As Date
''    Dim Lng_Year As Long
''    Dim Lng_Month As Long
''
''    Dim DB_1 As Database
''
''    Dim DS_1 As New ADODB.Recordset
''
''    Dim OP_FLG As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_This_YM
''
''    DoEvents
''
''    Ret = Fnc_Query_Open_ADO("QS11_This_Year", DS_1, OP_FLG)
''
''    DoEvents
''
''    If DS_1.EOF = False Then
''        DS_1.MoveFirst
''        DoEvents
''        If DS_1.EOF = False Then
''        ''Do While DS_1.EOF = False
''        ''Loop
''            DoEvents
''            Lng_Year = DS_1![Yea]
''            Lng_Month = DS_1![Mon]
''            DoEvents
''        Else
''            ''�e�[�u������擾�ł��Ȃ����́A���ݓ�������擾
''            Lng_Year = Fnc_Get_This_Year()
''            Lng_Month = Fnc_Get_This_Month()
''        End If
''    End If
''
''    If Fnc_Query_Close_ADO(DS_1, OP_FLG) = False Then
''        GoTo Exit_Fnc_Get_Table_This_YM
''    End If
''
''    Fnc_Get_Table_This_YM = DateSerial(Lng_Year, Lng_Month, 1)
''
''Exit_Fnc_Get_Table_This_YM:
''
''    On Error GoTo 0
''
''    If OP_FLG > 0 Then
''        If OP_FLG >= 1 Then
''            Ret = Fnc_Query_Close_ADO(DS_1, OP_FLG)
''        End If
''    End If
''
''    Exit Function
''
''Err_Fnc_Get_Table_This_YM:
''
''    Select Case Err
''        Case -2147467259
''            DoEvents
''            Resume
''        Case 3021
''            DoEvents
''            Resume
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''            Resume
''            Resume Next
''    End Select
''
''End Function

''Public Function Fnc_Get_Table_This_Year() As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_This_Year
''
''    DoEvents
''
''    Fnc_Get_Table_This_Year = Year(Fnc_Get_Table_This_YM())
''
''    DoEvents
''
''Exit_Fnc_Get_Table_This_Year:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_This_Year:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_This_Month() As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_This_Month
''
''    DoEvents
''
''    Fnc_Get_Table_This_Month = Month(Fnc_Get_Table_This_YM())
''
''    DoEvents
''
''Exit_Fnc_Get_Table_This_Month:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_This_Month:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_This_Year_Fmt() As String
''
''    On Error GoTo Err_Fnc_Get_Table_This_Year_Fmt
''
''    DoEvents
''
''    Fnc_Get_Table_This_Year_Fmt = Format(Fnc_Get_Table_This_Year(), "0000")
''
''    DoEvents
''
''Exit_Fnc_Get_Table_This_Year_Fmt:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_This_Year_Fmt:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_This_Month_Fmt() As String
''
''    On Error GoTo Err_Fnc_Get_Table_This_Month_Fmt
''
''    DoEvents
''
''    Fnc_Get_Table_This_Month_Fmt = Format(Fnc_Get_Table_This_Month(), "00")
''
''    DoEvents
''
''Exit_Fnc_Get_Table_This_Month_Fmt:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_This_Month_Fmt:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_Next_YM() As Date
''    Dim Lng_Year As Long
''    Dim Lng_Month As Long
''    Dim Dte_Wk As Date
''
''    Dim DB_1 As Database
''
''    Dim DS_1 As New ADODB.Recordset
''
''    Dim OP_FLG As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_Next_YM
''
''    DoEvents
''
''    Ret = Fnc_Query_Open_ADO("QS11_This_Year", DS_1, OP_FLG)
''
''    DoEvents
''
''    If DS_1.EOF = False Then
''        DS_1.MoveFirst
''        DoEvents
''        If DS_1.EOF = False Then
''        ''Do While DS_1.EOF = False
''        ''Loop
''            DoEvents
''            Lng_Year = DS_1![Yea_Nx]
''            Lng_Month = DS_1![Mon_Nx]
''            DoEvents
''        Else
''            ''�e�[�u������擾�ł��Ȃ����́A���ݓ�������v�Z
''            Lng_Year = Fnc_Get_This_Year()
''            Lng_Month = Fnc_Get_This_Month()
''            Dte_Wk = Fnc_Get_Calc_Next_Month(Lng_Year, Lng_Month)
''            Lng_Year = Year(Dte_Wk)
''            Lng_Month = Month(Dte_Wk)
''        End If
''    End If
''
''    If Fnc_Query_Close_ADO(DS_1, OP_FLG) = False Then
''        GoTo Exit_Fnc_Get_Table_Next_YM
''    End If
''
''    Fnc_Get_Table_Next_YM = DateSerial(Lng_Year, Lng_Month, 1)
''
''Exit_Fnc_Get_Table_Next_YM:
''
''    On Error GoTo 0
''
''    If OP_FLG > 0 Then
''        If OP_FLG >= 1 Then
''            Ret = Fnc_Query_Close_ADO(DS_1, OP_FLG)
''        End If
''    End If
''
''    Exit Function
''
''Err_Fnc_Get_Table_Next_YM:
''
''    Select Case Err
''        Case -2147467259
''            DoEvents
''            Resume
''        Case 3021
''            DoEvents
''            Resume
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''            Resume
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_Next_Year() As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_Next_Year
''
''    DoEvents
''
''    Fnc_Get_Table_Next_Year = Year(Fnc_Get_Table_Next_YM())
''
''    DoEvents
''
''Exit_Fnc_Get_Table_Next_Year:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_Next_Year:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_Next_Month() As Integer
''
''    On Error GoTo Err_Fnc_Get_Table_Next_Month:
''
''    DoEvents
''
''    Fnc_Get_Table_Next_Month = Month(Fnc_Get_Table_Next_YM())
''
''    DoEvents
''
''Exit_Fnc_Get_Table_Next_Month:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_Next_Month:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_Next_Year_Fmt() As String
''
''    On Error GoTo Err_Fnc_Get_Table_Next_Year_Fmt
''
''    DoEvents
''
''    Fnc_Get_Table_Next_Year_Fmt = Format(Fnc_Get_Table_Next_Year(), "0000")
''
''    DoEvents
''
''Exit_Fnc_Get_Table_Next_Year_Fmt:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_Next_Year_Fmt:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''
''Public Function Fnc_Get_Table_Next_Month_Fmt() As String
''
''    On Error GoTo Err_Fnc_Get_Table_Next_Month_Fmt
''
''    DoEvents
''
''    Fnc_Get_Table_Next_Month_Fmt = Format(Fnc_Get_Table_Next_Month(), "00")
''
''    DoEvents
''
''Exit_Fnc_Get_Table_Next_Month_Fmt:
''
''    On Error GoTo 0
''
''    Exit Function
''
''Err_Fnc_Get_Table_Next_Month_Fmt:
''
''    Select Case Err
''        Case Else                                                               '��L�ȊO�̃G���[
''            If DBEngine.Errors.Count > 0 Then
''                ' Errors �R���N�V������񋓂��܂��B
''                For Each Errloop In DBEngine.Errors
''                    MsgBox "Error number:" & Errloop.Number & _
''                        vbCr & Errloop.Description
''                Next Errloop
''            End If
''
''            Resume Next
''    End Select
''
''End Function
''2017/11/22 Add End

''2017/12/01 Add Start
Public Function Fnc_Mouse_Cur_Chg(Int_Chg_Mode As Integer) As Integer
    
    On Error GoTo Err_Fnc_Mouse_Cur_Chg
    
    DoEvents
    
    Select Case Int_Chg_Mode
        Case Con_Mouse_Wait
            DoCmd.Hourglass True
        Case Else
            DoCmd.Hourglass False
    End Select
    
    DoEvents

Exit_Fnc_Mouse_Cur_Chg:

    On Error GoTo 0

    Exit Function

Err_Fnc_Mouse_Cur_Chg:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Mouse_Cur_Wait() As Integer
    
    On Error GoTo Err_Fnc_Mouse_Cur_Wait
    
    DoEvents
    
    Fnc_Mouse_Cur_Wait = Fnc_Mouse_Cur_Chg(Con_Mouse_Wait)
    
    DoEvents

Exit_Fnc_Mouse_Cur_Wait:

    On Error GoTo 0

    Exit Function

Err_Fnc_Mouse_Cur_Wait:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

Public Function Fnc_Mouse_Cur_Nor() As Integer
    
    On Error GoTo Err_Fnc_Mouse_Cur_Nor
    
    DoEvents
    
    Fnc_Mouse_Cur_Nor = Fnc_Mouse_Cur_Chg(Con_Mouse_Nor)
    
    DoEvents

Exit_Fnc_Mouse_Cur_Nor:

    On Error GoTo 0

    Exit Function

Err_Fnc_Mouse_Cur_Nor:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function
''2017/12/01 Add End

'2017/12/06 Add Start
Public Function Fnc_Cmb_Year_Make(Obj_Combo As Object, Int_Base_Year As Integer) As Integer
    Dim Int_Start_Year As Integer
    Dim Int_Cnt As Integer
    
    Dim Int_Year_Range As Integer
    
    On Error GoTo Err_Fnc_Cmb_Year_Make
    
    Fnc_Cmb_Year_Make = False
    
    ''�O��T�N�ɐݒ�
    Int_Year_Range = 5
    
    Int_Start_Year = Int_Base_Year - Int_Year_Range
    
    For Int_Cnt = Int_Start_Year To Int_Base_Year + Int_Year_Range  '�O�� Int_Year_Range �N���̔N��o�^
        Obj_Combo.AddItem CStr(Int_Cnt)
        DoEvents
    Next Int_Cnt

    Fnc_Cmb_Year_Make = True

Exit_Fnc_Cmb_Year_Make:

    On Error GoTo 0

    Exit Function

Err_Fnc_Cmb_Year_Make:
    Select Case Err
'        Case -2147467259
'            DoEvents
'            Resume
'        Case 3021
'            DoEvents
'            Resume
            ''6014
            ''3001
        
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function
'2017/12/16 Add End

''2017/12/09 Add Start
Public Function Fnc_Form_Open(Str_Form_Name As String) As Integer
    
    Fnc_Form_Open = False
    
    On Error GoTo Err_Fnc_Form_Open
    
    DoEvents
    
    DoCmd.OpenForm Str_Form_Name, acNormal, , , acFormEdit, acWindowNormal
    
    DoEvents

    Fnc_Form_Open = True

Exit_Fnc_Form_Open:

    On Error GoTo 0
    
    Exit Function

Err_Fnc_Form_Open:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Form_Close(Str_Form_Name As String) As Integer

    On Error GoTo Err_Fnc_Form_Close

    Fnc_Form_Close = False

    DoEvents
    
    DoCmd.Close acForm, Str_Form_Name, acSaveNo
    
    DoEvents
    
    Fnc_Form_Close = True

Exit_Fnc_Form_Close:

    On Error GoTo 0

    Exit Function

Err_Fnc_Form_Close:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Message_Dsp(Str_Message As String) As Integer
    
    On Error GoTo Err_Fnc_Message_Dsp

    Fnc_Message_Dsp = False
    
    DoEvents
    
    Ret = Fnc_Form_Open("FS00_Message")

    ''2019/05/23 Add Start
    Ret = Fnc_DebugPrintFile("Fnc_Message_Dsp:" & Str_Message, "M00_Public_Module")
    ''2019/05/23 Add End
    
    DoEvents
    
    [Forms]![FS00_Message]![Txt_Message01] = Str_Message
    
    DoEvents
    
    [Forms]![FS00_Message]![Txt_Message02] = Str_Message
    
    DoEvents

    [Forms]![FS00_Message]![Txt_Sys_Msg] = ""
    
    DoEvents

    Fnc_Message_Dsp = True

Exit_Fnc_Message_Dsp:

    On Error GoTo 0

    Exit Function
    
Err_Fnc_Message_Dsp:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Message_Close() As Integer

    On Error GoTo Err_Fnc_Message_Close

    Fnc_Message_Close = False

    DoEvents
    
    Ret = Fnc_Form_Close("FS00_Message")
    
    DoEvents

    Fnc_Message_Close = True

Exit_Fnc_Message_Close:

    On Error GoTo 0

    Exit Function

Err_Fnc_Message_Close:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2017/12/09 Add End

''2017/12/11 Add Start
Public Function Fnc_Wait_Message_Dsp(Optional Str_Message As String) As Integer
    
    Dim Str_Wk_Message As String
    
    On Error GoTo Err_Fnc_Wait_Message_Dsp

    Fnc_Wait_Message_Dsp = False
    
    DoEvents
    
    Ret = Fnc_Mouse_Cur_Wait()
    
    DoEvents
    
    If Len(Trim(Str_Message)) = 0 Then
        Str_Wk_Message = Con_Proc_Wait_Msg
    Else
        Str_Wk_Message = Str_Message
    End If
    
    DoEvents
    
    Ret = Fnc_Message_Dsp(Str_Wk_Message)
    
    DoEvents

    Fnc_Wait_Message_Dsp = True

Exit_Fnc_Wait_Message_Dsp:

    On Error GoTo 0

    Exit Function

Err_Fnc_Wait_Message_Dsp:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Wait_Message_Close() As Integer

    On Error GoTo Err_Fnc_Wait_Message_Close
        
    Fnc_Wait_Message_Close = False

    DoEvents
    
    Ret = Fnc_Message_Close()
    
    DoEvents
    
    Ret = Fnc_Mouse_Cur_Nor()
    
    DoEvents

    Fnc_Wait_Message_Close = True

Exit_Fnc_Wait_Message_Close:

    On Error GoTo 0

    Exit Function

Err_Fnc_Wait_Message_Close:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2017/12/11 Add End
''2017/12/12 Add Start
Public Function Fnc_Sys_Msg_Dsp(Optional Str_Message As String) As Integer
    
    On Error GoTo Err_Fnc_Sys_Msg_Dsp

    Fnc_Sys_Msg_Dsp = False
    
    DoEvents

    ''2019/05/23 Add Start
    Ret = Fnc_DebugPrintFile("Fnc_Sys_Msg_Dsp:" & Str_Message, "M00_Public_Module")
    
    Ret = Fnc_Form_Open("FS00_Message")
    
    DoEvents

    [Forms]![FS00_Message]![Txt_Sys_Msg] = Str_Message
    
    DoEvents

    Fnc_Sys_Msg_Dsp = True

Exit_Fnc_Sys_Msg_Dsp:

    On Error GoTo 0

    Exit Function
    
Err_Fnc_Sys_Msg_Dsp:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2017/12/12 Add End

''2017/12/14 Add Start
Public Function Fnc_Required_Amount_Calc() As Integer

    Dim Lng_Calc_Wk(4, 65) As Long
    Dim Lng_Day_Cnt As Long
    Dim Lng_Day_Serch As Long
    Dim Lng_Day_Serch_Start As Long

    Dim DB_1 As Database

    Dim DB_2 As Database

    Dim DS_1 As New ADODB.Recordset
    Dim DS_2 As New ADODB.Recordset

    Dim OP_FLG As Integer

    Dim Lng_Data_Cnt As Long

    Dim Str_Wk_Date As String
    Dim Str_Wk_Suu As String

    On Error GoTo Err_Fnc_Required_Amount_Calc

    DoEvents

    Ret = Fnc_Query_Open_ADO("TD04_Material_Plan", DS_1, OP_FLG)

    DoEvents

    If DS_1.EOF = False Then

        DS_1.MoveFirst

        DoEvents

        Do While DS_1.EOF = False

            Lng_Data_Cnt = 0

            DoEvents

            ''�ǂݍ��݃f�[�^����U�e�[�u���Ɋi�[
            Ret = Fnc_Required_Amount_Data_Get(Lng_Calc_Wk, DS_1)

            Lng_Day_Serch_Start = 1

            Select Case DS_1![ProductNo_Key]
                Case "MTK0003"
                    Ret = Ret
                Case "MTK0005"
                    Ret = Ret
                Case Else

            End Select

            For Lng_Day_Cnt = 1 To UBound(Lng_Calc_Wk, 2)
                If Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) > 0 Then
''                    If Lng_Day_Cnt < Lng_Day_Serch_Start Then
''                        Lng_Day_Serch_Start = Lng_Day_Cnt
''                    End If
                    ''For Lng_Day_Serch = UBound(Lng_Calc_Wk, 2) To Lng_Day_Serch_Start Step -1
                    For Lng_Day_Serch = Lng_Day_Serch_Start To UBound(Lng_Calc_Wk, 2)
                        ''�K�v�ʂ��o�^����Ă��邩�H
                        If Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch) > 0 Then
                            ''���ѐ����K�v����菭�Ȃ����H
                            If Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch) <= Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) Then
                                ''�c���v�Z
                                Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) = Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) - Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch)
                                Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch) = 0
                            Else
                                Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch) = Lng_Calc_Wk(Con_Required_Amount_Point, Lng_Day_Serch) - Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt)
                                Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) = 0
                            End If

                            '�����ʒu���o����i�������x�A�b�v�j
                            Lng_Day_Serch_Start = Lng_Day_Serch

                            ''�c���L��H
                            If Lng_Calc_Wk(Con_Performance_Point, Lng_Day_Cnt) <= 0 Then
                                Exit For
                            End If
                        End If
                        If Lng_Day_Serch = 31 Then
                            Lng_Day_Serch = Lng_Day_Serch
                        End If
                    Next Lng_Day_Serch
                End If
            Next Lng_Day_Cnt

            ''�ҏW�σf�[�^�����R�[�h�Ɋi�[
            Ret = Fnc_Required_Amount_Data_Set(Lng_Calc_Wk, DS_1)

            DS_1.Update

            DoEvents

            DS_1.MoveNext

            DoEvents

        Loop
    End If

    ''If Fnc_Query_Close(DB_1, DS_1, OP_Flg) = False Then
    If Fnc_Query_Close_ADO(DS_1, OP_FLG) = False Then
        GoTo Exit_Fnc_Required_Amount_Calc
    End If

Exit_Fnc_Required_Amount_Calc:

    On Error GoTo 0

    If OP_FLG > 0 Then
        If OP_FLG >= 1 Then
            ''Ret = Fnc_Query_Close(DB_1, DS_1, OP_Flg)
            Ret = Fnc_Query_Close_ADO(DS_1, OP_FLG)
        End If
    End If

    Exit Function

Err_Fnc_Required_Amount_Calc:

    Select Case Err
        Case -2147467259
            DoEvents
            Resume
        Case 3021
            DoEvents
            Resume
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Private Function Fnc_Required_Amount_Data_Get(ByRef Lng_Calc_Wk() As Long, ByRef DS As Object) As Integer
    
    With DS
        ''2019/07/23 Add Start
        Lng_Calc_Wk(Con_Order_Point, 1) = ![Orders_Before]
        Lng_Calc_Wk(Con_Order_Point, 2) = ![Orders26_01]
        Lng_Calc_Wk(Con_Order_Point, 3) = ![Orders27_01]
        Lng_Calc_Wk(Con_Order_Point, 4) = ![Orders28_01]
        Lng_Calc_Wk(Con_Order_Point, 5) = ![Orders29_01]
        Lng_Calc_Wk(Con_Order_Point, 6) = ![Orders30_01]
        Lng_Calc_Wk(Con_Order_Point, 7) = ![Orders31_01]
        Lng_Calc_Wk(Con_Order_Point, 8) = ![Orders01_01]
        Lng_Calc_Wk(Con_Order_Point, 9) = ![Orders02_01]
        Lng_Calc_Wk(Con_Order_Point, 10) = ![Orders03_01]
        Lng_Calc_Wk(Con_Order_Point, 11) = ![Orders04_01]
        Lng_Calc_Wk(Con_Order_Point, 12) = ![Orders05_01]
        Lng_Calc_Wk(Con_Order_Point, 13) = ![Orders06_01]
        Lng_Calc_Wk(Con_Order_Point, 14) = ![Orders07_01]
        Lng_Calc_Wk(Con_Order_Point, 15) = ![Orders08_01]
        Lng_Calc_Wk(Con_Order_Point, 16) = ![Orders09_01]
        Lng_Calc_Wk(Con_Order_Point, 17) = ![Orders10_01]
        Lng_Calc_Wk(Con_Order_Point, 18) = ![Orders11_01]
        Lng_Calc_Wk(Con_Order_Point, 19) = ![Orders12_01]
        Lng_Calc_Wk(Con_Order_Point, 20) = ![Orders13_01]
        Lng_Calc_Wk(Con_Order_Point, 21) = ![Orders14_01]
        Lng_Calc_Wk(Con_Order_Point, 22) = ![Orders15_01]
        Lng_Calc_Wk(Con_Order_Point, 23) = ![Orders16_01]
        Lng_Calc_Wk(Con_Order_Point, 24) = ![Orders17_01]
        Lng_Calc_Wk(Con_Order_Point, 25) = ![Orders18_01]
        Lng_Calc_Wk(Con_Order_Point, 26) = ![Orders19_01]
        Lng_Calc_Wk(Con_Order_Point, 27) = ![Orders20_01]
        Lng_Calc_Wk(Con_Order_Point, 28) = ![Orders21_01]
        Lng_Calc_Wk(Con_Order_Point, 29) = ![Orders22_01]
        Lng_Calc_Wk(Con_Order_Point, 30) = ![Orders23_01]
        Lng_Calc_Wk(Con_Order_Point, 31) = ![Orders24_01]
        Lng_Calc_Wk(Con_Order_Point, 32) = ![Orders25_01]
        Lng_Calc_Wk(Con_Order_Point, 33) = ![Orders26_02]
        Lng_Calc_Wk(Con_Order_Point, 34) = ![Orders27_02]
        Lng_Calc_Wk(Con_Order_Point, 35) = ![Orders28_02]
        Lng_Calc_Wk(Con_Order_Point, 36) = ![Orders29_02]
        Lng_Calc_Wk(Con_Order_Point, 37) = ![Orders30_02]
        Lng_Calc_Wk(Con_Order_Point, 38) = ![Orders31_02]
        Lng_Calc_Wk(Con_Order_Point, 39) = ![Orders01_02]
        Lng_Calc_Wk(Con_Order_Point, 40) = ![Orders02_02]
        Lng_Calc_Wk(Con_Order_Point, 41) = ![Orders03_02]
        Lng_Calc_Wk(Con_Order_Point, 42) = ![Orders04_02]
        Lng_Calc_Wk(Con_Order_Point, 43) = ![Orders05_02]
        Lng_Calc_Wk(Con_Order_Point, 44) = ![Orders06_02]
        Lng_Calc_Wk(Con_Order_Point, 45) = ![Orders07_02]
        Lng_Calc_Wk(Con_Order_Point, 46) = ![Orders08_02]
        Lng_Calc_Wk(Con_Order_Point, 47) = ![Orders09_02]
        Lng_Calc_Wk(Con_Order_Point, 48) = ![Orders10_02]
        Lng_Calc_Wk(Con_Order_Point, 49) = ![Orders11_02]
        Lng_Calc_Wk(Con_Order_Point, 50) = ![Orders12_02]
        Lng_Calc_Wk(Con_Order_Point, 51) = ![Orders13_02]
        Lng_Calc_Wk(Con_Order_Point, 52) = ![Orders14_02]
        Lng_Calc_Wk(Con_Order_Point, 53) = ![Orders15_02]
        Lng_Calc_Wk(Con_Order_Point, 54) = ![Orders16_02]
        Lng_Calc_Wk(Con_Order_Point, 55) = ![Orders17_02]
        Lng_Calc_Wk(Con_Order_Point, 56) = ![Orders18_02]
        Lng_Calc_Wk(Con_Order_Point, 57) = ![Orders19_02]
        Lng_Calc_Wk(Con_Order_Point, 58) = ![Orders20_02]
        Lng_Calc_Wk(Con_Order_Point, 59) = ![Orders21_02]
        Lng_Calc_Wk(Con_Order_Point, 60) = ![Orders22_02]
        Lng_Calc_Wk(Con_Order_Point, 61) = ![Orders23_02]
        Lng_Calc_Wk(Con_Order_Point, 62) = ![Orders24_02]
        Lng_Calc_Wk(Con_Order_Point, 63) = ![Orders25_02]

        Lng_Calc_Wk(Con_Performance_Point, 1) = ![Performance_Before]
        Lng_Calc_Wk(Con_Performance_Point, 2) = ![Performance26_01]
        Lng_Calc_Wk(Con_Performance_Point, 3) = ![Performance27_01]
        Lng_Calc_Wk(Con_Performance_Point, 4) = ![Performance28_01]
        Lng_Calc_Wk(Con_Performance_Point, 5) = ![Performance29_01]
        Lng_Calc_Wk(Con_Performance_Point, 6) = ![Performance30_01]
        Lng_Calc_Wk(Con_Performance_Point, 7) = ![Performance31_01]
        Lng_Calc_Wk(Con_Performance_Point, 8) = ![Performance01_01]
        Lng_Calc_Wk(Con_Performance_Point, 9) = ![Performance02_01]
        Lng_Calc_Wk(Con_Performance_Point, 10) = ![Performance03_01]
        Lng_Calc_Wk(Con_Performance_Point, 11) = ![Performance04_01]
        Lng_Calc_Wk(Con_Performance_Point, 12) = ![Performance05_01]
        Lng_Calc_Wk(Con_Performance_Point, 13) = ![Performance06_01]
        Lng_Calc_Wk(Con_Performance_Point, 14) = ![Performance07_01]
        Lng_Calc_Wk(Con_Performance_Point, 15) = ![Performance08_01]
        Lng_Calc_Wk(Con_Performance_Point, 16) = ![Performance09_01]
        Lng_Calc_Wk(Con_Performance_Point, 17) = ![Performance10_01]
        Lng_Calc_Wk(Con_Performance_Point, 18) = ![Performance11_01]
        Lng_Calc_Wk(Con_Performance_Point, 19) = ![Performance12_01]
        Lng_Calc_Wk(Con_Performance_Point, 20) = ![Performance13_01]
        Lng_Calc_Wk(Con_Performance_Point, 21) = ![Performance14_01]
        Lng_Calc_Wk(Con_Performance_Point, 22) = ![Performance15_01]
        Lng_Calc_Wk(Con_Performance_Point, 23) = ![Performance16_01]
        Lng_Calc_Wk(Con_Performance_Point, 24) = ![Performance17_01]
        Lng_Calc_Wk(Con_Performance_Point, 25) = ![Performance18_01]
        Lng_Calc_Wk(Con_Performance_Point, 26) = ![Performance19_01]
        Lng_Calc_Wk(Con_Performance_Point, 27) = ![Performance20_01]
        Lng_Calc_Wk(Con_Performance_Point, 28) = ![Performance21_01]
        Lng_Calc_Wk(Con_Performance_Point, 29) = ![Performance22_01]
        Lng_Calc_Wk(Con_Performance_Point, 30) = ![Performance23_01]
        Lng_Calc_Wk(Con_Performance_Point, 31) = ![Performance24_01]
        Lng_Calc_Wk(Con_Performance_Point, 32) = ![Performance25_01]
        Lng_Calc_Wk(Con_Performance_Point, 33) = ![Performance26_02]
        Lng_Calc_Wk(Con_Performance_Point, 34) = ![Performance27_02]
        Lng_Calc_Wk(Con_Performance_Point, 35) = ![Performance28_02]
        Lng_Calc_Wk(Con_Performance_Point, 36) = ![Performance29_02]
        Lng_Calc_Wk(Con_Performance_Point, 37) = ![Performance30_02]
        Lng_Calc_Wk(Con_Performance_Point, 38) = ![Performance31_02]
        Lng_Calc_Wk(Con_Performance_Point, 39) = ![Performance01_02]
        Lng_Calc_Wk(Con_Performance_Point, 40) = ![Performance02_02]
        Lng_Calc_Wk(Con_Performance_Point, 41) = ![Performance03_02]
        Lng_Calc_Wk(Con_Performance_Point, 42) = ![Performance04_02]
        Lng_Calc_Wk(Con_Performance_Point, 43) = ![Performance05_02]
        Lng_Calc_Wk(Con_Performance_Point, 44) = ![Performance06_02]
        Lng_Calc_Wk(Con_Performance_Point, 45) = ![Performance07_02]
        Lng_Calc_Wk(Con_Performance_Point, 46) = ![Performance08_02]
        Lng_Calc_Wk(Con_Performance_Point, 47) = ![Performance09_02]
        Lng_Calc_Wk(Con_Performance_Point, 48) = ![Performance10_02]
        Lng_Calc_Wk(Con_Performance_Point, 49) = ![Performance11_02]
        Lng_Calc_Wk(Con_Performance_Point, 50) = ![Performance12_02]
        Lng_Calc_Wk(Con_Performance_Point, 51) = ![Performance13_02]
        Lng_Calc_Wk(Con_Performance_Point, 52) = ![Performance14_02]
        Lng_Calc_Wk(Con_Performance_Point, 53) = ![Performance15_02]
        Lng_Calc_Wk(Con_Performance_Point, 54) = ![Performance16_02]
        Lng_Calc_Wk(Con_Performance_Point, 55) = ![Performance17_02]
        Lng_Calc_Wk(Con_Performance_Point, 56) = ![Performance18_02]
        Lng_Calc_Wk(Con_Performance_Point, 57) = ![Performance19_02]
        Lng_Calc_Wk(Con_Performance_Point, 58) = ![Performance20_02]
        Lng_Calc_Wk(Con_Performance_Point, 59) = ![Performance21_02]
        Lng_Calc_Wk(Con_Performance_Point, 60) = ![Performance22_02]
        Lng_Calc_Wk(Con_Performance_Point, 61) = ![Performance23_02]
        Lng_Calc_Wk(Con_Performance_Point, 62) = ![Performance24_02]
        Lng_Calc_Wk(Con_Performance_Point, 63) = ![Performance25_02]

        Lng_Calc_Wk(Con_Delivery_Record_Point, 1) = ![Delivery_Record_Before]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 2) = ![Delivery_Record26_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 3) = ![Delivery_Record27_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 4) = ![Delivery_Record28_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 5) = ![Delivery_Record29_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 6) = ![Delivery_Record30_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 7) = ![Delivery_Record31_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 8) = ![Delivery_Record01_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 9) = ![Delivery_Record02_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 10) = ![Delivery_Record03_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 11) = ![Delivery_Record04_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 12) = ![Delivery_Record05_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 13) = ![Delivery_Record06_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 14) = ![Delivery_Record07_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 15) = ![Delivery_Record08_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 16) = ![Delivery_Record09_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 17) = ![Delivery_Record10_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 18) = ![Delivery_Record11_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 19) = ![Delivery_Record12_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 20) = ![Delivery_Record13_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 21) = ![Delivery_Record14_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 22) = ![Delivery_Record15_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 23) = ![Delivery_Record16_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 24) = ![Delivery_Record17_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 25) = ![Delivery_Record18_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 26) = ![Delivery_Record19_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 27) = ![Delivery_Record20_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 28) = ![Delivery_Record21_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 29) = ![Delivery_Record22_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 30) = ![Delivery_Record23_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 31) = ![Delivery_Record24_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 32) = ![Delivery_Record25_01]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 33) = ![Delivery_Record26_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 34) = ![Delivery_Record27_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 35) = ![Delivery_Record28_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 36) = ![Delivery_Record29_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 37) = ![Delivery_Record30_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 38) = ![Delivery_Record31_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 39) = ![Delivery_Record01_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 40) = ![Delivery_Record02_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 41) = ![Delivery_Record03_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 42) = ![Delivery_Record04_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 43) = ![Delivery_Record05_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 44) = ![Delivery_Record06_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 45) = ![Delivery_Record07_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 46) = ![Delivery_Record08_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 47) = ![Delivery_Record09_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 48) = ![Delivery_Record10_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 49) = ![Delivery_Record11_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 50) = ![Delivery_Record12_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 51) = ![Delivery_Record13_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 52) = ![Delivery_Record14_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 53) = ![Delivery_Record15_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 54) = ![Delivery_Record16_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 55) = ![Delivery_Record17_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 56) = ![Delivery_Record18_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 57) = ![Delivery_Record19_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 58) = ![Delivery_Record20_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 59) = ![Delivery_Record21_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 60) = ![Delivery_Record22_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 61) = ![Delivery_Record23_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 62) = ![Delivery_Record24_02]
        Lng_Calc_Wk(Con_Delivery_Record_Point, 63) = ![Delivery_Record25_02]

        ''Lng_Calc_Wk(Con_Required_Amount_Point, 1) = 0
        Lng_Calc_Wk(Con_Required_Amount_Point, 1) = ![Orders_Before]
        Lng_Calc_Wk(Con_Required_Amount_Point, 2) = ![Required_Amount26_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 3) = ![Required_Amount27_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 4) = ![Required_Amount28_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 5) = ![Required_Amount29_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 6) = ![Required_Amount30_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 7) = ![Required_Amount31_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 8) = ![Required_Amount01_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 9) = ![Required_Amount02_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 10) = ![Required_Amount03_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 11) = ![Required_Amount04_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 12) = ![Required_Amount05_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 13) = ![Required_Amount06_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 14) = ![Required_Amount07_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 15) = ![Required_Amount08_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 16) = ![Required_Amount09_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 17) = ![Required_Amount10_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 18) = ![Required_Amount11_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 19) = ![Required_Amount12_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 20) = ![Required_Amount13_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 21) = ![Required_Amount14_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 22) = ![Required_Amount15_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 23) = ![Required_Amount16_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 24) = ![Required_Amount17_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 25) = ![Required_Amount18_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 26) = ![Required_Amount19_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 27) = ![Required_Amount20_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 28) = ![Required_Amount21_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 29) = ![Required_Amount22_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 30) = ![Required_Amount23_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 31) = ![Required_Amount24_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 32) = ![Required_Amount25_01]
        Lng_Calc_Wk(Con_Required_Amount_Point, 33) = ![Required_Amount26_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 34) = ![Required_Amount27_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 35) = ![Required_Amount28_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 36) = ![Required_Amount29_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 37) = ![Required_Amount30_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 38) = ![Required_Amount31_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 39) = ![Required_Amount01_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 40) = ![Required_Amount02_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 41) = ![Required_Amount03_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 42) = ![Required_Amount04_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 43) = ![Required_Amount05_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 44) = ![Required_Amount06_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 45) = ![Required_Amount07_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 46) = ![Required_Amount08_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 47) = ![Required_Amount09_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 48) = ![Required_Amount10_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 49) = ![Required_Amount11_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 50) = ![Required_Amount12_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 51) = ![Required_Amount13_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 52) = ![Required_Amount14_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 53) = ![Required_Amount15_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 54) = ![Required_Amount16_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 55) = ![Required_Amount17_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 56) = ![Required_Amount18_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 57) = ![Required_Amount19_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 58) = ![Required_Amount20_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 59) = ![Required_Amount21_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 60) = ![Required_Amount22_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 61) = ![Required_Amount23_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 62) = ![Required_Amount24_02]
        Lng_Calc_Wk(Con_Required_Amount_Point, 63) = ![Required_Amount25_02]
        ''2019/07/23 Add End

        ''2019/07/23 Delete Start
''        Lng_Calc_Wk(Con_Order_Point, 1) = ![Orders_Before]
''        Lng_Calc_Wk(Con_Order_Point, 2) = ![Orders21_01]
''        Lng_Calc_Wk(Con_Order_Point, 3) = ![Orders22_01]
''        Lng_Calc_Wk(Con_Order_Point, 4) = ![Orders23_01]
''        Lng_Calc_Wk(Con_Order_Point, 5) = ![Orders24_01]
''        Lng_Calc_Wk(Con_Order_Point, 6) = ![Orders25_01]
''        Lng_Calc_Wk(Con_Order_Point, 7) = ![Orders26_01]
''        Lng_Calc_Wk(Con_Order_Point, 8) = ![Orders27_01]
''        Lng_Calc_Wk(Con_Order_Point, 9) = ![Orders28_01]
''        Lng_Calc_Wk(Con_Order_Point, 10) = ![Orders29_01]
''        Lng_Calc_Wk(Con_Order_Point, 11) = ![Orders30_01]
''        Lng_Calc_Wk(Con_Order_Point, 12) = ![Orders31_01]
''        Lng_Calc_Wk(Con_Order_Point, 13) = ![Orders01_01]
''        Lng_Calc_Wk(Con_Order_Point, 14) = ![Orders02_01]
''        Lng_Calc_Wk(Con_Order_Point, 15) = ![Orders03_01]
''        Lng_Calc_Wk(Con_Order_Point, 16) = ![Orders04_01]
''        Lng_Calc_Wk(Con_Order_Point, 17) = ![Orders05_01]
''        Lng_Calc_Wk(Con_Order_Point, 18) = ![Orders06_01]
''        Lng_Calc_Wk(Con_Order_Point, 19) = ![Orders07_01]
''        Lng_Calc_Wk(Con_Order_Point, 20) = ![Orders08_01]
''        Lng_Calc_Wk(Con_Order_Point, 21) = ![Orders09_01]
''        Lng_Calc_Wk(Con_Order_Point, 22) = ![Orders10_01]
''        Lng_Calc_Wk(Con_Order_Point, 23) = ![Orders11_01]
''        Lng_Calc_Wk(Con_Order_Point, 24) = ![Orders12_01]
''        Lng_Calc_Wk(Con_Order_Point, 25) = ![Orders13_01]
''        Lng_Calc_Wk(Con_Order_Point, 26) = ![Orders14_01]
''        Lng_Calc_Wk(Con_Order_Point, 27) = ![Orders15_01]
''        Lng_Calc_Wk(Con_Order_Point, 28) = ![Orders16_01]
''        Lng_Calc_Wk(Con_Order_Point, 29) = ![Orders17_01]
''        Lng_Calc_Wk(Con_Order_Point, 30) = ![Orders18_01]
''        Lng_Calc_Wk(Con_Order_Point, 31) = ![Orders19_01]
''        Lng_Calc_Wk(Con_Order_Point, 32) = ![Orders20_01]
''        Lng_Calc_Wk(Con_Order_Point, 33) = ![Orders21_02]
''        Lng_Calc_Wk(Con_Order_Point, 34) = ![Orders22_02]
''        Lng_Calc_Wk(Con_Order_Point, 35) = ![Orders23_02]
''        Lng_Calc_Wk(Con_Order_Point, 36) = ![Orders24_02]
''        Lng_Calc_Wk(Con_Order_Point, 37) = ![Orders25_02]
''        Lng_Calc_Wk(Con_Order_Point, 38) = ![Orders26_02]
''        Lng_Calc_Wk(Con_Order_Point, 39) = ![Orders27_02]
''        Lng_Calc_Wk(Con_Order_Point, 40) = ![Orders28_02]
''        Lng_Calc_Wk(Con_Order_Point, 41) = ![Orders29_02]
''        Lng_Calc_Wk(Con_Order_Point, 42) = ![Orders30_02]
''        Lng_Calc_Wk(Con_Order_Point, 43) = ![Orders31_02]
''        Lng_Calc_Wk(Con_Order_Point, 44) = ![Orders01_02]
''        Lng_Calc_Wk(Con_Order_Point, 45) = ![Orders02_02]
''        Lng_Calc_Wk(Con_Order_Point, 46) = ![Orders03_02]
''        Lng_Calc_Wk(Con_Order_Point, 47) = ![Orders04_02]
''        Lng_Calc_Wk(Con_Order_Point, 48) = ![Orders05_02]
''        Lng_Calc_Wk(Con_Order_Point, 49) = ![Orders06_02]
''        Lng_Calc_Wk(Con_Order_Point, 50) = ![Orders07_02]
''        Lng_Calc_Wk(Con_Order_Point, 51) = ![Orders08_02]
''        Lng_Calc_Wk(Con_Order_Point, 52) = ![Orders09_02]
''        Lng_Calc_Wk(Con_Order_Point, 53) = ![Orders10_02]
''        Lng_Calc_Wk(Con_Order_Point, 54) = ![Orders11_02]
''        Lng_Calc_Wk(Con_Order_Point, 55) = ![Orders12_02]
''        Lng_Calc_Wk(Con_Order_Point, 56) = ![Orders13_02]
''        Lng_Calc_Wk(Con_Order_Point, 57) = ![Orders14_02]
''        Lng_Calc_Wk(Con_Order_Point, 58) = ![Orders15_02]
''        Lng_Calc_Wk(Con_Order_Point, 59) = ![Orders16_02]
''        Lng_Calc_Wk(Con_Order_Point, 60) = ![Orders17_02]
''        Lng_Calc_Wk(Con_Order_Point, 61) = ![Orders18_02]
''        Lng_Calc_Wk(Con_Order_Point, 62) = ![Orders19_02]
''        Lng_Calc_Wk(Con_Order_Point, 63) = ![Orders20_02]
''
''        Lng_Calc_Wk(Con_Performance_Point, 1) = ![Performance_Before]
''        Lng_Calc_Wk(Con_Performance_Point, 2) = ![Performance21_01]
''        Lng_Calc_Wk(Con_Performance_Point, 3) = ![Performance22_01]
''        Lng_Calc_Wk(Con_Performance_Point, 4) = ![Performance23_01]
''        Lng_Calc_Wk(Con_Performance_Point, 5) = ![Performance24_01]
''        Lng_Calc_Wk(Con_Performance_Point, 6) = ![Performance25_01]
''        Lng_Calc_Wk(Con_Performance_Point, 7) = ![Performance26_01]
''        Lng_Calc_Wk(Con_Performance_Point, 8) = ![Performance27_01]
''        Lng_Calc_Wk(Con_Performance_Point, 9) = ![Performance28_01]
''        Lng_Calc_Wk(Con_Performance_Point, 10) = ![Performance29_01]
''        Lng_Calc_Wk(Con_Performance_Point, 11) = ![Performance30_01]
''        Lng_Calc_Wk(Con_Performance_Point, 12) = ![Performance31_01]
''        Lng_Calc_Wk(Con_Performance_Point, 13) = ![Performance01_01]
''        Lng_Calc_Wk(Con_Performance_Point, 14) = ![Performance02_01]
''        Lng_Calc_Wk(Con_Performance_Point, 15) = ![Performance03_01]
''        Lng_Calc_Wk(Con_Performance_Point, 16) = ![Performance04_01]
''        Lng_Calc_Wk(Con_Performance_Point, 17) = ![Performance05_01]
''        Lng_Calc_Wk(Con_Performance_Point, 18) = ![Performance06_01]
''        Lng_Calc_Wk(Con_Performance_Point, 19) = ![Performance07_01]
''        Lng_Calc_Wk(Con_Performance_Point, 20) = ![Performance08_01]
''        Lng_Calc_Wk(Con_Performance_Point, 21) = ![Performance09_01]
''        Lng_Calc_Wk(Con_Performance_Point, 22) = ![Performance10_01]
''        Lng_Calc_Wk(Con_Performance_Point, 23) = ![Performance11_01]
''        Lng_Calc_Wk(Con_Performance_Point, 24) = ![Performance12_01]
''        Lng_Calc_Wk(Con_Performance_Point, 25) = ![Performance13_01]
''        Lng_Calc_Wk(Con_Performance_Point, 26) = ![Performance14_01]
''        Lng_Calc_Wk(Con_Performance_Point, 27) = ![Performance15_01]
''        Lng_Calc_Wk(Con_Performance_Point, 28) = ![Performance16_01]
''        Lng_Calc_Wk(Con_Performance_Point, 29) = ![Performance17_01]
''        Lng_Calc_Wk(Con_Performance_Point, 30) = ![Performance18_01]
''        Lng_Calc_Wk(Con_Performance_Point, 31) = ![Performance19_01]
''        Lng_Calc_Wk(Con_Performance_Point, 32) = ![Performance20_01]
''        Lng_Calc_Wk(Con_Performance_Point, 33) = ![Performance21_02]
''        Lng_Calc_Wk(Con_Performance_Point, 34) = ![Performance22_02]
''        Lng_Calc_Wk(Con_Performance_Point, 35) = ![Performance23_02]
''        Lng_Calc_Wk(Con_Performance_Point, 36) = ![Performance24_02]
''        Lng_Calc_Wk(Con_Performance_Point, 37) = ![Performance25_02]
''        Lng_Calc_Wk(Con_Performance_Point, 38) = ![Performance26_02]
''        Lng_Calc_Wk(Con_Performance_Point, 39) = ![Performance27_02]
''        Lng_Calc_Wk(Con_Performance_Point, 40) = ![Performance28_02]
''        Lng_Calc_Wk(Con_Performance_Point, 41) = ![Performance29_02]
''        Lng_Calc_Wk(Con_Performance_Point, 42) = ![Performance30_02]
''        Lng_Calc_Wk(Con_Performance_Point, 43) = ![Performance31_02]
''        Lng_Calc_Wk(Con_Performance_Point, 44) = ![Performance01_02]
''        Lng_Calc_Wk(Con_Performance_Point, 45) = ![Performance02_02]
''        Lng_Calc_Wk(Con_Performance_Point, 46) = ![Performance03_02]
''        Lng_Calc_Wk(Con_Performance_Point, 47) = ![Performance04_02]
''        Lng_Calc_Wk(Con_Performance_Point, 48) = ![Performance05_02]
''        Lng_Calc_Wk(Con_Performance_Point, 49) = ![Performance06_02]
''        Lng_Calc_Wk(Con_Performance_Point, 50) = ![Performance07_02]
''        Lng_Calc_Wk(Con_Performance_Point, 51) = ![Performance08_02]
''        Lng_Calc_Wk(Con_Performance_Point, 52) = ![Performance09_02]
''        Lng_Calc_Wk(Con_Performance_Point, 53) = ![Performance10_02]
''        Lng_Calc_Wk(Con_Performance_Point, 54) = ![Performance11_02]
''        Lng_Calc_Wk(Con_Performance_Point, 55) = ![Performance12_02]
''        Lng_Calc_Wk(Con_Performance_Point, 56) = ![Performance13_02]
''        Lng_Calc_Wk(Con_Performance_Point, 57) = ![Performance14_02]
''        Lng_Calc_Wk(Con_Performance_Point, 58) = ![Performance15_02]
''        Lng_Calc_Wk(Con_Performance_Point, 59) = ![Performance16_02]
''        Lng_Calc_Wk(Con_Performance_Point, 60) = ![Performance17_02]
''        Lng_Calc_Wk(Con_Performance_Point, 61) = ![Performance18_02]
''        Lng_Calc_Wk(Con_Performance_Point, 62) = ![Performance19_02]
''        Lng_Calc_Wk(Con_Performance_Point, 63) = ![Performance20_02]
''
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 1) = ![Delivery_Record_Before]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 2) = ![Delivery_Record21_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 3) = ![Delivery_Record22_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 4) = ![Delivery_Record23_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 5) = ![Delivery_Record24_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 6) = ![Delivery_Record25_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 7) = ![Delivery_Record26_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 8) = ![Delivery_Record27_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 9) = ![Delivery_Record28_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 10) = ![Delivery_Record29_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 11) = ![Delivery_Record30_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 12) = ![Delivery_Record31_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 13) = ![Delivery_Record01_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 14) = ![Delivery_Record02_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 15) = ![Delivery_Record03_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 16) = ![Delivery_Record04_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 17) = ![Delivery_Record05_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 18) = ![Delivery_Record06_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 19) = ![Delivery_Record07_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 20) = ![Delivery_Record08_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 21) = ![Delivery_Record09_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 22) = ![Delivery_Record10_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 23) = ![Delivery_Record11_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 24) = ![Delivery_Record12_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 25) = ![Delivery_Record13_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 26) = ![Delivery_Record14_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 27) = ![Delivery_Record15_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 28) = ![Delivery_Record16_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 29) = ![Delivery_Record17_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 30) = ![Delivery_Record18_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 31) = ![Delivery_Record19_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 32) = ![Delivery_Record20_01]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 33) = ![Delivery_Record21_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 34) = ![Delivery_Record22_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 35) = ![Delivery_Record23_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 36) = ![Delivery_Record24_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 37) = ![Delivery_Record25_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 38) = ![Delivery_Record26_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 39) = ![Delivery_Record27_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 40) = ![Delivery_Record28_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 41) = ![Delivery_Record29_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 42) = ![Delivery_Record30_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 43) = ![Delivery_Record31_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 44) = ![Delivery_Record01_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 45) = ![Delivery_Record02_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 46) = ![Delivery_Record03_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 47) = ![Delivery_Record04_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 48) = ![Delivery_Record05_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 49) = ![Delivery_Record06_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 50) = ![Delivery_Record07_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 51) = ![Delivery_Record08_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 52) = ![Delivery_Record09_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 53) = ![Delivery_Record10_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 54) = ![Delivery_Record11_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 55) = ![Delivery_Record12_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 56) = ![Delivery_Record13_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 57) = ![Delivery_Record14_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 58) = ![Delivery_Record15_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 59) = ![Delivery_Record16_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 60) = ![Delivery_Record17_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 61) = ![Delivery_Record18_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 62) = ![Delivery_Record19_02]
''        Lng_Calc_Wk(Con_Delivery_Record_Point, 63) = ![Delivery_Record20_02]
''
''        ''Lng_Calc_Wk(Con_Required_Amount_Point, 1) = 0
''        Lng_Calc_Wk(Con_Required_Amount_Point, 1) = ![Orders_Before]
''
''        Lng_Calc_Wk(Con_Required_Amount_Point, 2) = ![Required_Amount21_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 3) = ![Required_Amount22_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 4) = ![Required_Amount23_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 5) = ![Required_Amount24_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 6) = ![Required_Amount25_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 7) = ![Required_Amount26_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 8) = ![Required_Amount27_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 9) = ![Required_Amount28_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 10) = ![Required_Amount29_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 11) = ![Required_Amount30_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 12) = ![Required_Amount31_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 13) = ![Required_Amount01_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 14) = ![Required_Amount02_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 15) = ![Required_Amount03_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 16) = ![Required_Amount04_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 17) = ![Required_Amount05_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 18) = ![Required_Amount06_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 19) = ![Required_Amount07_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 20) = ![Required_Amount08_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 21) = ![Required_Amount09_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 22) = ![Required_Amount10_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 23) = ![Required_Amount11_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 24) = ![Required_Amount12_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 25) = ![Required_Amount13_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 26) = ![Required_Amount14_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 27) = ![Required_Amount15_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 28) = ![Required_Amount16_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 29) = ![Required_Amount17_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 30) = ![Required_Amount18_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 31) = ![Required_Amount19_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 32) = ![Required_Amount20_01]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 33) = ![Required_Amount21_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 34) = ![Required_Amount22_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 35) = ![Required_Amount23_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 36) = ![Required_Amount24_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 37) = ![Required_Amount25_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 38) = ![Required_Amount26_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 39) = ![Required_Amount27_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 40) = ![Required_Amount28_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 41) = ![Required_Amount29_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 42) = ![Required_Amount30_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 43) = ![Required_Amount31_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 44) = ![Required_Amount01_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 45) = ![Required_Amount02_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 46) = ![Required_Amount03_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 47) = ![Required_Amount04_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 48) = ![Required_Amount05_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 49) = ![Required_Amount06_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 50) = ![Required_Amount07_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 51) = ![Required_Amount08_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 52) = ![Required_Amount09_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 53) = ![Required_Amount10_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 54) = ![Required_Amount11_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 55) = ![Required_Amount12_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 56) = ![Required_Amount13_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 57) = ![Required_Amount14_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 58) = ![Required_Amount15_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 59) = ![Required_Amount16_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 60) = ![Required_Amount17_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 61) = ![Required_Amount18_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 62) = ![Required_Amount19_02]
''        Lng_Calc_Wk(Con_Required_Amount_Point, 63) = ![Required_Amount20_02]
        ''2019/07/23 Delete End
    
    End With
End Function

Private Function Fnc_Required_Amount_Data_Set(ByRef Lng_Calc_Wk() As Long, ByRef DS As Object) As Integer
    With DS
''        ![Orders_Before] = Lng_Calc_Wk(Con_Order_Point, 1)
''        ![Orders21_01] = Lng_Calc_Wk(Con_Order_Point, 2)
''        ![Orders22_01] = Lng_Calc_Wk(Con_Order_Point, 3)
''        ![Orders23_01] = Lng_Calc_Wk(Con_Order_Point, 4)
''        ![Orders24_01] = Lng_Calc_Wk(Con_Order_Point, 5)
''        ![Orders25_01] = Lng_Calc_Wk(Con_Order_Point, 6)
''        ![Orders26_01] = Lng_Calc_Wk(Con_Order_Point, 7)
''        ![Orders27_01] = Lng_Calc_Wk(Con_Order_Point, 8)
''        ![Orders28_01] = Lng_Calc_Wk(Con_Order_Point, 9)
''        ![Orders29_01] = Lng_Calc_Wk(Con_Order_Point, 10)
''        ![Orders30_01] = Lng_Calc_Wk(Con_Order_Point, 11)
''        ![Orders31_01] = Lng_Calc_Wk(Con_Order_Point, 12)
''        ![Orders01_01] = Lng_Calc_Wk(Con_Order_Point, 13)
''        ![Orders02_01] = Lng_Calc_Wk(Con_Order_Point, 14)
''        ![Orders03_01] = Lng_Calc_Wk(Con_Order_Point, 15)
''        ![Orders04_01] = Lng_Calc_Wk(Con_Order_Point, 16)
''        ![Orders05_01] = Lng_Calc_Wk(Con_Order_Point, 17)
''        ![Orders06_01] = Lng_Calc_Wk(Con_Order_Point, 18)
''        ![Orders07_01] = Lng_Calc_Wk(Con_Order_Point, 19)
''        ![Orders08_01] = Lng_Calc_Wk(Con_Order_Point, 20)
''        ![Orders09_01] = Lng_Calc_Wk(Con_Order_Point, 21)
''        ![Orders10_01] = Lng_Calc_Wk(Con_Order_Point, 22)
''        ![Orders11_01] = Lng_Calc_Wk(Con_Order_Point, 23)
''        ![Orders12_01] = Lng_Calc_Wk(Con_Order_Point, 24)
''        ![Orders13_01] = Lng_Calc_Wk(Con_Order_Point, 25)
''        ![Orders14_01] = Lng_Calc_Wk(Con_Order_Point, 26)
''        ![Orders15_01] = Lng_Calc_Wk(Con_Order_Point, 27)
''        ![Orders16_01] = Lng_Calc_Wk(Con_Order_Point, 28)
''        ![Orders17_01] = Lng_Calc_Wk(Con_Order_Point, 29)
''        ![Orders18_01] = Lng_Calc_Wk(Con_Order_Point, 30)
''        ![Orders19_01] = Lng_Calc_Wk(Con_Order_Point, 31)
''        ![Orders20_01] = Lng_Calc_Wk(Con_Order_Point, 32)
''        ![Orders21_02] = Lng_Calc_Wk(Con_Order_Point, 33)
''        ![Orders22_02] = Lng_Calc_Wk(Con_Order_Point, 34)
''        ![Orders23_02] = Lng_Calc_Wk(Con_Order_Point, 35)
''        ![Orders24_02] = Lng_Calc_Wk(Con_Order_Point, 36)
''        ![Orders25_02] = Lng_Calc_Wk(Con_Order_Point, 37)
''        ![Orders26_02] = Lng_Calc_Wk(Con_Order_Point, 38)
''        ![Orders27_02] = Lng_Calc_Wk(Con_Order_Point, 39)
''        ![Orders28_02] = Lng_Calc_Wk(Con_Order_Point, 40)
''        ![Orders29_02] = Lng_Calc_Wk(Con_Order_Point, 41)
''        ![Orders30_02] = Lng_Calc_Wk(Con_Order_Point, 42)
''        ![Orders31_02] = Lng_Calc_Wk(Con_Order_Point, 43)
''        ![Orders01_02] = Lng_Calc_Wk(Con_Order_Point, 44)
''        ![Orders02_02] = Lng_Calc_Wk(Con_Order_Point, 45)
''        ![Orders03_02] = Lng_Calc_Wk(Con_Order_Point, 46)
''        ![Orders04_02] = Lng_Calc_Wk(Con_Order_Point, 47)
''        ![Orders05_02] = Lng_Calc_Wk(Con_Order_Point, 48)
''        ![Orders06_02] = Lng_Calc_Wk(Con_Order_Point, 49)
''        ![Orders07_02] = Lng_Calc_Wk(Con_Order_Point, 50)
''        ![Orders08_02] = Lng_Calc_Wk(Con_Order_Point, 51)
''        ![Orders09_02] = Lng_Calc_Wk(Con_Order_Point, 52)
''        ![Orders10_02] = Lng_Calc_Wk(Con_Order_Point, 53)
''        ![Orders11_02] = Lng_Calc_Wk(Con_Order_Point, 54)
''        ![Orders12_02] = Lng_Calc_Wk(Con_Order_Point, 55)
''        ![Orders13_02] = Lng_Calc_Wk(Con_Order_Point, 56)
''        ![Orders14_02] = Lng_Calc_Wk(Con_Order_Point, 57)
''        ![Orders15_02] = Lng_Calc_Wk(Con_Order_Point, 58)
''        ![Orders16_02] = Lng_Calc_Wk(Con_Order_Point, 59)
''        ![Orders17_02] = Lng_Calc_Wk(Con_Order_Point, 60)
''        ![Orders18_02] = Lng_Calc_Wk(Con_Order_Point, 61)
''        ![Orders19_02] = Lng_Calc_Wk(Con_Order_Point, 62)
''        ![Orders20_02] = Lng_Calc_Wk(Con_Order_Point, 63)

''        ![Performance_Before] = Lng_Calc_Wk(Con_Performance_Point, 1)
''        ![Performance21_01] = Lng_Calc_Wk(Con_Performance_Point, 2)
''        ![Performance22_01] = Lng_Calc_Wk(Con_Performance_Point, 3)
''        ![Performance23_01] = Lng_Calc_Wk(Con_Performance_Point, 4)
''        ![Performance24_01] = Lng_Calc_Wk(Con_Performance_Point, 5)
''        ![Performance25_01] = Lng_Calc_Wk(Con_Performance_Point, 6)
''        ![Performance26_01] = Lng_Calc_Wk(Con_Performance_Point, 7)
''        ![Performance27_01] = Lng_Calc_Wk(Con_Performance_Point, 8)
''        ![Performance28_01] = Lng_Calc_Wk(Con_Performance_Point, 9)
''        ![Performance29_01] = Lng_Calc_Wk(Con_Performance_Point, 10)
''        ![Performance30_01] = Lng_Calc_Wk(Con_Performance_Point, 11)
''        ![Performance31_01] = Lng_Calc_Wk(Con_Performance_Point, 12)
''        ![Performance01_01] = Lng_Calc_Wk(Con_Performance_Point, 13)
''        ![Performance02_01] = Lng_Calc_Wk(Con_Performance_Point, 14)
''        ![Performance03_01] = Lng_Calc_Wk(Con_Performance_Point, 15)
''        ![Performance04_01] = Lng_Calc_Wk(Con_Performance_Point, 16)
''        ![Performance05_01] = Lng_Calc_Wk(Con_Performance_Point, 17)
''        ![Performance06_01] = Lng_Calc_Wk(Con_Performance_Point, 18)
''        ![Performance07_01] = Lng_Calc_Wk(Con_Performance_Point, 19)
''        ![Performance08_01] = Lng_Calc_Wk(Con_Performance_Point, 20)
''        ![Performance09_01] = Lng_Calc_Wk(Con_Performance_Point, 21)
''        ![Performance10_01] = Lng_Calc_Wk(Con_Performance_Point, 22)
''        ![Performance11_01] = Lng_Calc_Wk(Con_Performance_Point, 23)
''        ![Performance12_01] = Lng_Calc_Wk(Con_Performance_Point, 24)
''        ![Performance13_01] = Lng_Calc_Wk(Con_Performance_Point, 25)
''        ![Performance14_01] = Lng_Calc_Wk(Con_Performance_Point, 26)
''        ![Performance15_01] = Lng_Calc_Wk(Con_Performance_Point, 27)
''        ![Performance16_01] = Lng_Calc_Wk(Con_Performance_Point, 28)
''        ![Performance17_01] = Lng_Calc_Wk(Con_Performance_Point, 29)
''        ![Performance18_01] = Lng_Calc_Wk(Con_Performance_Point, 30)
''        ![Performance19_01] = Lng_Calc_Wk(Con_Performance_Point, 31)
''        ![Performance20_01] = Lng_Calc_Wk(Con_Performance_Point, 32)
''        ![Performance21_02] = Lng_Calc_Wk(Con_Performance_Point, 33)
''        ![Performance22_02] = Lng_Calc_Wk(Con_Performance_Point, 34)
''        ![Performance23_02] = Lng_Calc_Wk(Con_Performance_Point, 35)
''        ![Performance24_02] = Lng_Calc_Wk(Con_Performance_Point, 36)
''        ![Performance25_02] = Lng_Calc_Wk(Con_Performance_Point, 37)
''        ![Performance26_02] = Lng_Calc_Wk(Con_Performance_Point, 38)
''        ![Performance27_02] = Lng_Calc_Wk(Con_Performance_Point, 39)
''        ![Performance28_02] = Lng_Calc_Wk(Con_Performance_Point, 40)
''        ![Performance29_02] = Lng_Calc_Wk(Con_Performance_Point, 41)
''        ![Performance30_02] = Lng_Calc_Wk(Con_Performance_Point, 42)
''        ![Performance31_02] = Lng_Calc_Wk(Con_Performance_Point, 43)
''        ![Performance01_02] = Lng_Calc_Wk(Con_Performance_Point, 44)
''        ![Performance02_02] = Lng_Calc_Wk(Con_Performance_Point, 45)
''        ![Performance03_02] = Lng_Calc_Wk(Con_Performance_Point, 46)
''        ![Performance04_02] = Lng_Calc_Wk(Con_Performance_Point, 47)
''        ![Performance05_02] = Lng_Calc_Wk(Con_Performance_Point, 48)
''        ![Performance06_02] = Lng_Calc_Wk(Con_Performance_Point, 49)
''        ![Performance07_02] = Lng_Calc_Wk(Con_Performance_Point, 50)
''        ![Performance08_02] = Lng_Calc_Wk(Con_Performance_Point, 51)
''        ![Performance09_02] = Lng_Calc_Wk(Con_Performance_Point, 52)
''        ![Performance10_02] = Lng_Calc_Wk(Con_Performance_Point, 53)
''        ![Performance11_02] = Lng_Calc_Wk(Con_Performance_Point, 54)
''        ![Performance12_02] = Lng_Calc_Wk(Con_Performance_Point, 55)
''        ![Performance13_02] = Lng_Calc_Wk(Con_Performance_Point, 56)
''        ![Performance14_02] = Lng_Calc_Wk(Con_Performance_Point, 57)
''        ![Performance15_02] = Lng_Calc_Wk(Con_Performance_Point, 58)
''        ![Performance16_02] = Lng_Calc_Wk(Con_Performance_Point, 59)
''        ![Performance17_02] = Lng_Calc_Wk(Con_Performance_Point, 60)
''        ![Performance18_02] = Lng_Calc_Wk(Con_Performance_Point, 61)
''        ![Performance19_02] = Lng_Calc_Wk(Con_Performance_Point, 62)
''        ![Performance20_02] = Lng_Calc_Wk(Con_Performance_Point, 63)

''        ![Delivery_Record_Before] = Lng_Calc_Wk(Con_Delivery_Record_Point, 1)
''        ![Delivery_Record21_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 2)
''        ![Delivery_Record22_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 3)
''        ![Delivery_Record23_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 4)
''        ![Delivery_Record24_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 5)
''        ![Delivery_Record25_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 6)
''        ![Delivery_Record26_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 7)
''        ![Delivery_Record27_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 8)
''        ![Delivery_Record28_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 9)
''        ![Delivery_Record29_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 10)
''        ![Delivery_Record30_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 11)
''        ![Delivery_Record31_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 12)
''        ![Delivery_Record01_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 13)
''        ![Delivery_Record02_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 14)
''        ![Delivery_Record03_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 15)
''        ![Delivery_Record04_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 16)
''        ![Delivery_Record05_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 17)
''        ![Delivery_Record06_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 18)
''        ![Delivery_Record07_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 19)
''        ![Delivery_Record08_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 20)
''        ![Delivery_Record09_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 21)
''        ![Delivery_Record10_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 22)
''        ![Delivery_Record11_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 23)
''        ![Delivery_Record12_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 24)
''        ![Delivery_Record13_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 25)
''        ![Delivery_Record14_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 26)
''        ![Delivery_Record15_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 27)
''        ![Delivery_Record16_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 28)
''        ![Delivery_Record17_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 29)
''        ![Delivery_Record18_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 30)
''        ![Delivery_Record19_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 31)
''        ![Delivery_Record20_01] = Lng_Calc_Wk(Con_Delivery_Record_Point, 32)
''        ![Delivery_Record21_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 33)
''        ![Delivery_Record22_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 34)
''        ![Delivery_Record23_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 35)
''        ![Delivery_Record24_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 36)
''        ![Delivery_Record25_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 37)
''        ![Delivery_Record26_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 38)
''        ![Delivery_Record27_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 39)
''        ![Delivery_Record28_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 40)
''        ![Delivery_Record29_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 41)
''        ![Delivery_Record30_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 42)
''        ![Delivery_Record31_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 43)
''        ![Delivery_Record01_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 44)
''        ![Delivery_Record02_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 45)
''        ![Delivery_Record03_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 46)
''        ![Delivery_Record04_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 47)
''        ![Delivery_Record05_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 48)
''        ![Delivery_Record06_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 49)
''        ![Delivery_Record07_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 50)
''        ![Delivery_Record08_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 51)
''        ![Delivery_Record09_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 52)
''        ![Delivery_Record10_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 53)
''        ![Delivery_Record11_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 54)
''        ![Delivery_Record12_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 55)
''        ![Delivery_Record13_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 56)
''        ![Delivery_Record14_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 57)
''        ![Delivery_Record15_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 58)
''        ![Delivery_Record16_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 59)
''        ![Delivery_Record17_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 60)
''        ![Delivery_Record18_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 61)
''        ![Delivery_Record19_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 62)
''        ![Delivery_Record20_02] = Lng_Calc_Wk(Con_Delivery_Record_Point, 63)

        ''0   =   Lng_Calc_Wk(Con_Required_Amount_Point, 1)
        ''2019/07/23 Add Start
        ![Required_Amount26_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 2)
        ![Required_Amount27_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 3)
        ![Required_Amount28_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 4)
        ![Required_Amount29_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 5)
        ![Required_Amount30_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 6)
        ![Required_Amount31_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 7)
        ![Required_Amount01_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 8)
        ![Required_Amount02_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 9)
        ![Required_Amount03_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 10)
        ![Required_Amount04_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 11)
        ![Required_Amount05_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 12)
        ![Required_Amount06_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 13)
        ![Required_Amount07_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 14)
        ![Required_Amount08_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 15)
        ![Required_Amount09_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 16)
        ![Required_Amount10_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 17)
        ![Required_Amount11_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 18)
        ![Required_Amount12_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 19)
        ![Required_Amount13_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 20)
        ![Required_Amount14_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 21)
        ![Required_Amount15_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 22)
        ![Required_Amount16_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 23)
        ![Required_Amount17_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 24)
        ![Required_Amount18_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 25)
        ![Required_Amount19_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 26)
        ![Required_Amount20_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 27)
        ![Required_Amount21_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 28)
        ![Required_Amount22_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 29)
        ![Required_Amount23_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 30)
        ![Required_Amount24_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 31)
        ![Required_Amount25_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 32)
        ![Required_Amount26_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 33)
        ![Required_Amount27_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 34)
        ![Required_Amount28_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 35)
        ![Required_Amount29_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 36)
        ![Required_Amount30_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 37)
        ![Required_Amount31_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 38)
        ![Required_Amount01_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 39)
        ![Required_Amount02_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 40)
        ![Required_Amount03_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 41)
        ![Required_Amount04_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 42)
        ![Required_Amount05_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 43)
        ![Required_Amount06_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 44)
        ![Required_Amount07_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 45)
        ![Required_Amount08_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 46)
        ![Required_Amount09_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 47)
        ![Required_Amount10_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 48)
        ![Required_Amount11_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 49)
        ![Required_Amount12_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 50)
        ![Required_Amount13_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 51)
        ![Required_Amount14_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 52)
        ![Required_Amount15_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 53)
        ![Required_Amount16_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 54)
        ![Required_Amount17_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 55)
        ![Required_Amount18_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 56)
        ![Required_Amount19_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 57)
        ![Required_Amount20_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 58)
        ![Required_Amount21_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 59)
        ![Required_Amount22_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 60)
        ![Required_Amount23_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 61)
        ![Required_Amount24_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 62)
        ![Required_Amount25_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 63)
        ''2019/07/23 Add End

        ''2019/07/23 Delete Start
''        ![Required_Amount21_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 2)
''        ![Required_Amount22_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 3)
''        ![Required_Amount23_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 4)
''        ![Required_Amount24_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 5)
''        ![Required_Amount25_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 6)
''        ![Required_Amount26_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 7)
''        ![Required_Amount27_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 8)
''        ![Required_Amount28_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 9)
''        ![Required_Amount29_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 10)
''        ![Required_Amount30_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 11)
''        ![Required_Amount31_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 12)
''        ![Required_Amount01_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 13)
''        ![Required_Amount02_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 14)
''        ![Required_Amount03_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 15)
''        ![Required_Amount04_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 16)
''        ![Required_Amount05_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 17)
''        ![Required_Amount06_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 18)
''        ![Required_Amount07_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 19)
''        ![Required_Amount08_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 20)
''        ![Required_Amount09_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 21)
''        ![Required_Amount10_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 22)
''        ![Required_Amount11_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 23)
''        ![Required_Amount12_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 24)
''        ![Required_Amount13_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 25)
''        ![Required_Amount14_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 26)
''        ![Required_Amount15_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 27)
''        ![Required_Amount16_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 28)
''        ![Required_Amount17_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 29)
''        ![Required_Amount18_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 30)
''        ![Required_Amount19_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 31)
''        ![Required_Amount20_01] = Lng_Calc_Wk(Con_Required_Amount_Point, 32)
''        ![Required_Amount21_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 33)
''        ![Required_Amount22_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 34)
''        ![Required_Amount23_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 35)
''        ![Required_Amount24_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 36)
''        ![Required_Amount25_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 37)
''        ![Required_Amount26_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 38)
''        ![Required_Amount27_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 39)
''        ![Required_Amount28_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 40)
''        ![Required_Amount29_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 41)
''        ![Required_Amount30_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 42)
''        ![Required_Amount31_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 43)
''        ![Required_Amount01_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 44)
''        ![Required_Amount02_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 45)
''        ![Required_Amount03_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 46)
''        ![Required_Amount04_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 47)
''        ![Required_Amount05_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 48)
''        ![Required_Amount06_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 49)
''        ![Required_Amount07_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 50)
''        ![Required_Amount08_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 51)
''        ![Required_Amount09_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 52)
''        ![Required_Amount10_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 53)
''        ![Required_Amount11_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 54)
''        ![Required_Amount12_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 55)
''        ![Required_Amount13_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 56)
''        ![Required_Amount14_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 57)
''        ![Required_Amount15_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 58)
''        ![Required_Amount16_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 59)
''        ![Required_Amount17_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 60)
''        ![Required_Amount18_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 61)
''        ![Required_Amount19_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 62)
''        ![Required_Amount20_02] = Lng_Calc_Wk(Con_Required_Amount_Point, 63)
        ''2019/07/23 Delete End
    
    
    End With
End Function

''2017/12/14 Add End

''2017/12/19 Add Start
Public Function Fnc_Product_Master_Mode_Chg(Int_Mode_Local As Integer) As Integer
    
    With Forms!FM01_02_Product_Master
        .Cmb_Proc_Mode = Int_Mode_Local

        Select Case Int_Mode_Local
            Case Con_Proc_Mode_None
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = True
                
                ''����
                .Cmd_Close.Enabled = True
    
                ''�ύX
                .Cmd_Update.Enabled = False
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = False
                
                Ret = Fnc_Product_Master_Detail_Chg(False)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Con_Proc_Mode_New
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg(True)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = True
            
            Case Con_Proc_Mode_CopyNew
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg(True)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = True
            
            Case Con_Proc_Mode_Update
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg(True)
                
                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Con_Proc_Mode_Delete
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg(True)
                
                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Else
        End Select
    End With

End Function

Public Function Fnc_Product_Master_Detail_Chg(Int_Mode_Local As Integer) As Integer
    With Forms!FM01_02_Product_Master
        .Frm_Detail.Visible = Int_Mode_Local

        ''�L�[
        .Lbl_ProductNo_Key.Visible = Int_Mode_Local
        .Txt_ProductNo_Key.Visible = Int_Mode_Local

        ''���iNo
        .Lbl_ProductNo.Visible = Int_Mode_Local
        .Txt_ProductNo.Visible = Int_Mode_Local

        ''�[����
        .Lbl_Supplier_Code.Visible = Int_Mode_Local
        .Cmb_Supplier_Code.Visible = Int_Mode_Local

        ''�ޗ��R�[�h
        .Lbl_TUBE_MATERIAL_CODE.Visible = Int_Mode_Local
        .Cmb_Tube_Material_Code.Visible = Int_Mode_Local
        .Lbl_Supplier_Name.Visible = Int_Mode_Local

        ''�ޗ��ڍ�
        .Lbl_Material_Inf.Visible = Int_Mode_Local
        .Frm_Material_Inf.Visible = Int_Mode_Local
    
        ''�ގ�
        .Lbl_Material_Detail.Visible = Int_Mode_Local
        .Txt_Material_Detail.Visible = Int_Mode_Local
    
        ''�O�a
        .Lbl_OuterD.Visible = Int_Mode_Local
        .Txt_OuterD.Visible = Int_Mode_Local

        ''��
        .Lbl_PlateThickness.Visible = Int_Mode_Local
        .Txt_PlateThickness.Visible = Int_Mode_Local

        ''����
        .Lbl_LongMaterial.Visible = Int_Mode_Local
        .Txt_LongMaterial.Visible = Int_Mode_Local

        ''���ޏd��
        .Lbl_LongMaterialWeight.Visible = Int_Mode_Local
        .Txt_LongMaterialWeight.Visible = Int_Mode_Local

        ''�ؒf��
        .Lbl_Length.Visible = Int_Mode_Local
        .Txt_Length.Visible = Int_Mode_Local
        
        ''�P�d��(kg)
        .Lbl_SingleWeight.Visible = Int_Mode_Local
        .Txt_SingleWeight.Visible = Int_Mode_Local

        ''�؎̂�
        .Lbl_Truncation.Visible = Int_Mode_Local
        .Txt_Truncation.Visible = Int_Mode_Local

        ''�搔
        .Lbl_Participants.Visible = Int_Mode_Local
        .Txt_Participants.Visible = Int_Mode_Local

        ''�g�p��
        .Lbl_UseRate.Visible = Int_Mode_Local
        .Txt_UseRate.Visible = Int_Mode_Local

        ''�P��
        .Lbl_UnitPrice.Visible = Int_Mode_Local
        .Txt_UnitPrice.Visible = Int_Mode_Local

        ''�o�^����
        .Lbl_RegistDate.Visible = Int_Mode_Local
        .Txt_RegistDate.Visible = Int_Mode_Local

        ''�X�V����
        .Lbl_UpdateDate.Visible = Int_Mode_Local
        .Txt_UpdateDate.Visible = Int_Mode_Local

        ''���l
        .Lbl_MEMO.Visible = Int_Mode_Local
        .Txt_MEMO.Visible = Int_Mode_Local

        Select Case Int_Mode_Local
            Case True
                .Pic_Arrow01.Visible = True
                .Pic_Arrow02.Visible = False
            Case False
                .Pic_Arrow01.Visible = False
                .Pic_Arrow02.Visible = True
            Case Else
        End Select
    
    End With
End Function
''2017/12/19 Add End

''2017/12/22 Add Start
Public Function Fnc_Tube_Material_Mode_Chg(Int_Mode_Local As Integer) As Integer
    
    With Forms!FM03_01_Tube_Material
        .Cmb_Proc_Mode = Int_Mode_Local

        Select Case Int_Mode_Local
            Case Con_Proc_Mode_None
                
                ''�V�K
                .Cmd_New.Enabled = True
                
                ''����
                .Cmd_Close.Enabled = True
    
                ''�ύX
                .Cmd_Update.Enabled = False
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = False
                
                Ret = Fnc_Tube_Material_Detail_Chg(False)
                
                With Forms!FM03_01_Tube_Material
                    ''�L�[
                    .Txt_TubeMaterial_Code.Enabled = False
                End With
            
            Case Con_Proc_Mode_New
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Tube_Material_Detail_Chg(True)
            
                With Forms!FM03_01_Tube_Material
                    ''�L�[
                    .Txt_TubeMaterial_Code.Enabled = False
                End With
            
            Case Con_Proc_Mode_CopyNew
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Tube_Material_Detail_Chg(True)
            
                With Forms!FM03_01_Tube_Material
                    ''�L�[
                    .Txt_TubeMaterial_Code.Enabled = False
                End With
            Case Con_Proc_Mode_Update
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Tube_Material_Detail_Chg(True)
                With Forms!FM03_01_Tube_Material
                    ''�L�[
                    .Txt_TubeMaterial_Code.Enabled = False
                End With
            Case Con_Proc_Mode_Delete
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Tube_Material_Detail_Chg(True)
                
                With Forms!FM03_01_Tube_Material
                    ''�L�[
                    .Txt_TubeMaterial_Code.Enabled = False
                End With
            
            Case Else
        End Select
    End With

End Function

Public Function Fnc_Tube_Material_Detail_Chg(Int_Mode_Local As Integer) As Integer
    With Forms!FM03_01_Tube_Material
        .Frm_Detail.Visible = Int_Mode_Local

        ''�L�[
        .Lbl_TubeMaterial_Code.Visible = Int_Mode_Local
        .Txt_TubeMaterial_Code.Visible = Int_Mode_Local

        ''�ގ��R�[�h
        .Lbl_Material_Code.Visible = Int_Mode_Local
        .Cmb_Material_Code.Visible = Int_Mode_Local
        .Txt_Material_Code.Visible = Int_Mode_Local

        ''�O�a
        .Lbl_OuterD.Visible = Int_Mode_Local
        .Txt_OuterD.Visible = Int_Mode_Local
        
        ''��
        .Lbl_PlateThickness.Visible = Int_Mode_Local
        .Txt_PlateThickness.Visible = Int_Mode_Local
    
        ''����
        .Lbl_LongMaterial.Visible = Int_Mode_Local
        .Txt_LongMaterial.Visible = Int_Mode_Local
    
        ''���ޏd��
        .Lbl_LongMaterialWeight.Visible = Int_Mode_Local
        .Txt_LongMaterialWeight.Visible = Int_Mode_Local
    
        ''�o�^����
        .Lbl_RegistDate.Visible = Int_Mode_Local
        .Txt_RegistDate.Visible = Int_Mode_Local
    
        ''�X�V����
        .Lbl_UpdateDate.Visible = Int_Mode_Local
        .Txt_UpdateDate.Visible = Int_Mode_Local

        ''���l
        .Lbl_MEMO.Visible = Int_Mode_Local
        .Txt_MEMO.Visible = Int_Mode_Local

        Select Case Int_Mode_Local
            Case True
                .Pic_Arrow01.Visible = True
                .Pic_Arrow02.Visible = False
            Case False
                .Pic_Arrow01.Visible = False
                .Pic_Arrow02.Visible = True
            Case Else
        End Select
    
    End With
End Function

Public Function Fnc_Tube_Material_Get() As String
'********************************************************************************
'*
'*  �ޗ��R�[�h�̎擾����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    On Error GoTo Err_Fnc_Tube_Material_Get

    DoEvents
    
    Fnc_Tube_Material_Get = DFirst("TubeMaterial_Code_New", "QS07_00_Tube_Material_New_Get")

    DoEvents

Exit_Fnc_Tube_Material_Get:

    On Error GoTo 0

    Exit Function

Err_Fnc_Tube_Material_Get:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Get_Record_Count_SQL(Str_SQL As String) As Long

    Fnc_Get_Record_Count_SQL = -1

    On Error GoTo Err_Fnc_Get_Record_Count_SQL

    '�� �����ݒ� ��
    Dim rst    As DAO.Recordset   '���R�[�h�Z�b�g

    '�쐬����SQL���Ń��R�[�h�Z�b�g�쐬
    Set rst = CurrentDb.OpenRecordset(Str_SQL)

    '���R�[�h�������擾
    ''Fnc_Get_Record_Count_SQL = Rst![Str_Count_Field]
    Fnc_Get_Record_Count_SQL = rst.RecordCount

    '�� �I������ ��
    rst.Close
    Set rst = Nothing

Exit_Fnc_Get_Record_Count_SQL:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Record_Count_SQL:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function
''2017/12/22 Add End

''2018/03/01 Add Start
Public Function Fnc_Query_Delete() As Integer
 Dim cn As New ADODB.Connection
 Dim cat As New ADOX.Catalog
 Dim vew As ADOX.View
 Dim DataFlag As Integer
 
 Set cn = New ADODB.Connection
 cn.ConnectionString = _
    "Provider=microsoft.jet.oledb.4.0;" & _
    "Data Source=D:\NorthWIND.MDB"
  cn.Open
 Set cat.ActiveConnection = cn
 '�N�G���̑��݃`�F�b�N
 For Each vew In cat.Views
  Select Case vew.Name
    Case "1995�N ���i�敪�ʔ��㍂"
      DataFlag = 1
   End Select
 Next vew
 '�N�G�������݂��Ă���ꍇ�͍폜
 If DataFlag = 1 Then
   cat.Views.Delete ("1995�N ���i�敪�ʔ��㍂")
 Else
   MsgBox "�N�G���u1995�N ���i�敪�ʔ��㍂�v�����݂��܂���"
   GoTo �I������
 End If

�I������:
 cn.Close
 Set cn = Nothing
 Set cat = Nothing

End Function
Public Function Fnc_Query_Create() As Integer
    
    Dim Qdf    As DAO.QueryDef
    Dim strSQL As String

    strSQL = "Select * From �e�[�u��1"

    'CreateQueryDef���\�b�h�ɂ��N�G���쐬
    Set Qdf = CurrentDb.CreateQueryDef("�V�K�N�G��_DAO1", strSQL)
    CurrentDb.QueryDefs.Refresh

    'Append���\�b�h�ɂ��N�G���쐬
    Set Qdf = New QueryDef
    Qdf.Name = "�V�K�N�G��_DAO2"
    Qdf.SQL = strSQL
    CurrentDb.QueryDefs.Append Qdf
    CurrentDb.QueryDefs.Refresh

    '�I������
    Set Qdf = Nothing

End Function
''2018/03/01 Add End

''2018/10/19 Add Start
Public Function Fnc_Proc_Wait(Optional ByVal Int_W_Time As Integer = 0) As Integer
'********************************************************************************
'*
'*  �����E���荞��
'*
'*-------------------------------------------------------------------------------
'*
'*   ����               :   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'

    Fnc_Proc_Wait = False

    On Error GoTo Err_Fnc_Proc_Wait

    DoEvents

    If Int_W_Time = 0 Then
        Ret = Fnc_Wait_Timer(Proc_Wait)
    Else
        Ret = Fnc_Wait_Timer(Int_W_Time)
    End If

    Fnc_Proc_Wait = True

Exit_Fnc_Proc_Wait:

    On Error GoTo 0

    Exit Function


Err_Fnc_Proc_Wait:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Wait_Timer(W_Time) As Integer
'********************************************************************************
'*
'*  �������f����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  W_Time  :   �҂�����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'
    Dim Start_Time

    Fnc_Wait_Timer = False

    On Error GoTo Err_Fnc_Wait_Timer

    Start_Time = Time

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents

    Do While Fnc_Time_Chk_Sec(Start_Time) < W_Time
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Loop

    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents

    Fnc_Wait_Timer = True

Exit_Fnc_Wait_Timer:

    On Error GoTo 0

    Exit Function


Err_Fnc_Wait_Timer:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Time_Chk(I_Time) As Date
'********************************************************************************
'*
'*  ���ԃ`�F�b�N����(���ԒP��)
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  I_Date  :   �J�n���t
'*                  I_Time  :   �J�n����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'

    Dim Wk_Time

    Fnc_Time_Chk = False

    On Error GoTo Err_Fnc_Time_Chk

    Wk_Time = Time
    
    Fnc_Time_Chk = Wk_Time - I_Time

Exit_Fnc_Time_Chk:

    On Error GoTo 0

    Exit Function


Err_Fnc_Time_Chk:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Time_Chk_Sec(I_Time) As Long
'********************************************************************************
'*
'*  ���ԃ`�F�b�N�����i�b�P�ʁj
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  I_Date  :   �J�n���t
'*                  I_Time  :   �J�n����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'

    Dim Wk_Time

    Fnc_Time_Chk_Sec = 0

    On Error GoTo Err_Fnc_Time_Chk_Sec

    Wk_Time = Time
    
    Fnc_Time_Chk_Sec = Second(Wk_Time - I_Time)

Exit_Fnc_Time_Chk_Sec:

    On Error GoTo 0

    Exit Function


Err_Fnc_Time_Chk_Sec:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function ''2018/10/19 Add End

''2018/12/21 Add Start
Public Function Fnc_Input_Chk_Num(Obj_Ctl As Object, Str_Chk_Name As String, Int_Chk_Type As Integer) As Integer
'********************************************************************************
'*
'*  ���͒l�`�F�b�N(���l)
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  Obj_Ctl         :   �`�F�b�N�E�I�u�W�F�N�g
'*                  Str_Chk_Name    :   �`�F�b�N�E�R���g���[���E���ږ�
'*                  Int_Chk_Type    :   �`�F�b�N�E�I�u�W�F�N�g���
'*                                      1   :   �e�L�X�g�E�{�b�N�X
'*                                      2   :   �R���{�E�{�b�N�X
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'
    Dim Str_Wk As String
    Dim Str_Wk_S As String
    Dim Lng_Cnt As Long
    Dim Lng_Cnt_E As Long
    Dim Lng_Cnt_DT As Long

    Fnc_Input_Chk_Num = False
       
    On Error GoTo Err_Fnc_Input_Chk_Num
    
    If Fnc_Null_Chk(Obj_Ctl) = True Then
    End If
    
    If Fnc_Null_Chk(Obj_Ctl) = True Or Trim(Obj_Ctl) = "" Then
        Select Case Int_Chk_Type
            Case 1
                MsgBox ("�y" & Str_Chk_Name & "�z����͂��ĉ������B")
            Case 2
                MsgBox ("�y" & Str_Chk_Name & "�z��I�����ĉ������B")
            Case Else
                MsgBox ("�y" & Str_Chk_Name & "�z��o�^���ĉ������B")
        End Select
        GoTo Exit_Fnc_Input_Chk_Num
    Else
        ''�O��X�y�[�X���Ȃ��āA������擾
        Str_Wk = Trim(Obj_Ctl)

        ''������Ԃ̔��p�X�y�[�X�E�`�F�b�N
        If InStr(Str_Wk, " ") > 0 Then
            ''������Ԃɔ��p�X�y�[�X�L
            MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�̊Ԃɔ��p�X�y�[�X���L��܂��B" & vbCrLf & "���p�X�y�[�X����菜���ĉ������B")
            GoTo Exit_Fnc_Input_Chk_Num
        Else
            ''������Ԃ̑S�p�X�y�[�X�E�`�F�b�N
            If InStr(Str_Wk, "�@") > 0 Then
                ''������ԂɑS�p�X�y�[�X�L
                MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�̊ԂɑS�p�X�y�[�X���L��܂��B" & vbCrLf & "�S�p�X�y�[�X����菜���ĉ������B")
                GoTo Exit_Fnc_Input_Chk_Num
            Else
                ''��������s�`�F�b�N
                If InStr(Str_Wk, vbLf) > 0 Then
                    ''������ԂɑS�p�X�y�[�X�L
                    MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�̉��s���܂܂�Ă��܂��B" & vbCrLf & "���s����菜���ĉ������B")
                    GoTo Exit_Fnc_Input_Chk_Num
                Else

                    ''�����_�J�E���g�E�N���A
                    Lng_Cnt_DT = 0

                    ''�����񒷂̎擾
                    Lng_Cnt_E = Len(Str_Wk)

                    For Lng_Cnt = 1 To Lng_Cnt_E
                        ''�P�����擾
                        Str_Wk_S = Mid(Str_Wk, Lng_Cnt, 1)
                        Select Case Str_Wk_S
                            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                            Case "."
                                Lng_Cnt_DT = Lng_Cnt_DT + 1
                            Case ","
                                MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�ɃJ���}�y,�z�����͂���Ă��܂��B" & vbCrLf & "�J���}���Ȃ��ē��͂��ĉ������B")
                                GoTo Exit_Fnc_Input_Chk_Num
                            Case Else
                                MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�ɐ��l�ȊO�̓��͕������܂܂�Ă��܂��B" & vbCrLf & "���������l����͂��ĉ������B")
                                GoTo Exit_Fnc_Input_Chk_Num
                        End Select
                    Next Lng_Cnt
                    If Lng_Cnt_DT > 1 Then
                        MsgBox ("�y" & Str_Chk_Name & "�z" & "���͒l�ɏ����_�y.�z���������͂���Ă��܂��B" & vbCrLf & "�s�v�ȏ��ѐ��_���Ȃ��ĉ������B")
                        GoTo Exit_Fnc_Input_Chk_Num
                    End If
                End If
            End If
        End If
    End If

    Fnc_Input_Chk_Num = True

Exit_Fnc_Input_Chk_Num:

    On Error GoTo 0

    Exit Function

Err_Fnc_Input_Chk_Num:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_Null_Chk(Obj_Data As Object) As Integer
'********************************************************************************
'*
'*  �m�������`�F�b�N�i�S�p�X�y�[�X�Ή��j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  Obj_Data        :   �`�F�b�N�E�I�u�W�F�N�g
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'

    Dim Str_Wk As String

    Fnc_Null_Chk = False

    On Error GoTo Err_Fnc_Null_Chk

    If Not (IsNull(Trim(Obj_Data)) = True Or Trim(Obj_Data) = "") Then
        ''�X�y�[�XNull�ϊ�
        Str_Wk = Fnc_SP_2_Null(Obj_Data)
        If Str_Wk <> "" Then
            Exit Function
        End If
    End If

    Fnc_Null_Chk = True

Exit_Fnc_Null_Chk:

    On Error GoTo 0

    Exit Function

Err_Fnc_Null_Chk:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

Public Function Fnc_SP_2_Null(Obj_Data As Object) As String
'********************************************************************************
'*
'*  �m�������ϊ��i�S�p�X�y�[�X�Ή��j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*                  Obj_Data        :   �`�F�b�N�E�I�u�W�F�N�g
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True    :   ����I��
'*
'*           False   :   �X�V��
'*
'********************************************************************************
'

    Dim Str_Wk As String

    Fnc_SP_2_Null = ""

    On Error GoTo Err_Fnc_SP_2_Null

    Str_Wk = Obj_Data

    ''���p�X�y�[�X��Null�ɒu��
    Str_Wk = Replace(Str_Wk, " ", "")

    ''�S�p�X�y�[�X��Null�ɒu��
    Str_Wk = Replace(Str_Wk, "�@", "")

    Fnc_SP_2_Null = Str_Wk

Exit_Fnc_SP_2_Null:

    On Error GoTo 0

    Exit Function

Err_Fnc_SP_2_Null:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function
''2018/12/21 Add End

''2018/12/26 Add Start
Public Function Fnc_Product_No_Sub_Get() As String
'********************************************************************************
'*
'*  �ؒf�i�ԍ��̎擾����
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*          True            �F  ����I��
'*          False           �F  �ُ�I��
'*
'********************************************************************************
'

    On Error GoTo Err_Fnc_Product_No_Sub_Get

    DoEvents
    
    ''--''
    Fnc_Product_No_Sub_Get = (DFirst("Make_Product_Sub_No", "QS02_TM04_03_Product_Sub_No_Make"))

    DoEvents

Exit_Fnc_Product_No_Sub_Get:

    On Error GoTo 0

    Exit Function

Err_Fnc_Product_No_Sub_Get:
    
    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function
''2018/12/26 Add End

''2019/01/09 Add Start
Public Function Fnc_Product_Master_Mode_Chg2(Int_Mode_Local As Integer) As Integer
    
    With Forms!FM01_03_Product_Master2
    ''With Forms!FM01_03_Product_MasterT
        .Cmb_Proc_Mode = Int_Mode_Local

        Select Case Int_Mode_Local
            Case Con_Proc_Mode_None
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = True
                
                ''����
                .Cmd_Close.Enabled = True
    
                ''�ύX
                .Cmd_Update.Enabled = False
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = False
                
                Ret = Fnc_Product_Master_Detail_Chg2(False)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Con_Proc_Mode_New
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg2(True)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = True
            
            Case Con_Proc_Mode_CopyNew
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = False
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg2(True)

                ''�L�[
                .Txt_ProductNo_Key.Enabled = True
            
            Case Con_Proc_Mode_Update
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg2(True)
                
                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Con_Proc_Mode_Delete
                ''�ޗ��ꗗ
                .Cmd_Material_Detail_List.Enabled = True
                
                ''���i�ꗗ
                .Cmd_Product_List.Enabled = True
                
                ''�V�K
                .Cmd_New.Enabled = False
                
                ''����
                .Cmd_Close.Enabled = False
    
                ''�ύX
                .Cmd_Update.Enabled = True
    
                ''�폜
                .Cmd_Delete.Enabled = True
                
                ''�L�����Z��
                .Cmd_Cancel.Enabled = True
                
                Ret = Fnc_Product_Master_Detail_Chg2(True)
                
                ''�L�[
                .Txt_ProductNo_Key.Enabled = False
            
            Case Else
        End Select
    End With

End Function

Public Function Fnc_Product_Master_Detail_Chg2(Int_Mode_Local As Integer) As Integer
    With Forms!FM01_03_Product_Master2
    ''With Forms!FM01_03_Product_MasterT
        .Frm_Detail.Visible = Int_Mode_Local

        ''�L�[
        .Lbl_ProductNo_Key.Visible = Int_Mode_Local
        .Txt_ProductNo_Key.Visible = Int_Mode_Local

        ''���iNo
        .Lbl_ProductNo.Visible = Int_Mode_Local
        .Txt_ProductNo.Visible = Int_Mode_Local

        ''�[����
        .Lbl_Supplier_Code.Visible = Int_Mode_Local
        .Cmb_Supplier_Code.Visible = Int_Mode_Local

        ''�ؒf�i
        .Lbl_PRODUCTNO_SUB_KEY.Visible = Int_Mode_Local
        .Cmb_PRODUCTNO_SUB_KEY.Visible = Int_Mode_Local

        ''�ޗ��R�[�h
        .Lbl_TUBE_MATERIAL_CODE.Visible = Int_Mode_Local
        .Cmb_Tube_Material_Code.Visible = Int_Mode_Local
        .Lbl_Supplier_Name.Visible = Int_Mode_Local

        ''�ޗ��ڍ�
        .Lbl_Material_Inf.Visible = Int_Mode_Local
        .Frm_Material_Inf.Visible = Int_Mode_Local
    
        ''�ގ�
        .Lbl_Material_Detail.Visible = Int_Mode_Local
        .Txt_Material_Detail.Visible = Int_Mode_Local
    
        ''�O�a
        .Lbl_OuterD.Visible = Int_Mode_Local
        .Txt_OuterD.Visible = Int_Mode_Local

        ''��
        .Lbl_PlateThickness.Visible = Int_Mode_Local
        .Txt_PlateThickness.Visible = Int_Mode_Local

        ''����
        .Lbl_LongMaterial.Visible = Int_Mode_Local
        .Txt_LongMaterial.Visible = Int_Mode_Local

        ''���ޏd��
        .Lbl_LongMaterialWeight.Visible = Int_Mode_Local
        .Txt_LongMaterialWeight.Visible = Int_Mode_Local
        
        ''�ؒf��
        .Lbl_Length.Visible = Int_Mode_Local
        .Txt_Length.Visible = Int_Mode_Local
        
        ''�P�d��(kg)
        .Lbl_SingleWeight.Visible = Int_Mode_Local
        .Txt_SingleWeight.Visible = Int_Mode_Local

        ''�؎̂�
        .Lbl_Truncation.Visible = Int_Mode_Local
        .Txt_Truncation.Visible = Int_Mode_Local

        ''�搔
        .Lbl_Participants.Visible = Int_Mode_Local
        .Txt_Participants.Visible = Int_Mode_Local

        ''�g�p��
        .Lbl_UseRate.Visible = Int_Mode_Local
        .Txt_UseRate.Visible = Int_Mode_Local

        ''�P��
        .Lbl_UnitPrice.Visible = Int_Mode_Local
        .Txt_UnitPrice.Visible = Int_Mode_Local

        ''�o�^����
        .Lbl_RegistDate.Visible = Int_Mode_Local
        .Txt_RegistDate.Visible = Int_Mode_Local

        ''�X�V����
        .Lbl_UpdateDate.Visible = Int_Mode_Local
        .Txt_UpdateDate.Visible = Int_Mode_Local

        ''���l
        .Lbl_MEMO.Visible = Int_Mode_Local
        .Txt_MEMO.Visible = Int_Mode_Local

        ''2019/03/18 Add Start
        ''�ؒf�i���
        .Lbl_Cut.Visible = Int_Mode_Local
        .Frm_Cut.Visible = Int_Mode_Local

        ''�t�C���^�E�t���[��
        .Lbl_Filter.Visible = Int_Mode_Local
        .Frm_Filter.Visible = Int_Mode_Local

        ''�t�C���^�E�t���[��
        ''�ޗ��i�ގ��E�O�a�E���j
        .Lbl_Cmb_Material_Fil.Visible = Int_Mode_Local
        .Cmb_Material_Fil.Visible = Int_Mode_Local
        ''2019/03/18 Add End

        Select Case Int_Mode_Local
            Case True
                .Pic_Arrow01.Visible = True
                .Pic_Arrow02.Visible = False
            Case False
                .Pic_Arrow01.Visible = False
                .Pic_Arrow02.Visible = True
            Case Else
        End Select
    
    End With
End Function
''2019/01/09 Add End

Public Function Fnc_Report_Preview(ByVal Str_Report_Name As String) As Integer

    On Error GoTo Err_Fnc_Report_Preview

    Fnc_Report_Preview = False

    DoCmd.OpenReport Str_Report_Name, acViewPreview

Exit_Fnc_Report_Preview:

    On Error GoTo 0

    Exit Function

Err_Fnc_Report_Preview:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume Next
    End Select

End Function

''2019/03/22 Add Start
Public Function Fnc_ADO_Open(Obj_Cn As Object, Obj_RS As Object, Str_SQL As String, OP_FLG As Integer) As Integer
'********************************************************************************
'*
'*  �N�G���[�E�I�[�v�������i�N�G���[���w��Ver�j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*          T_Name      :   �e�[�u����
'*          DB_Open     :   �f�[�^�E�x�[�X��`
'*          DS_Open     :   ���R�[�h�E�Z�b�g��`
'*          OP_Flg      :   �t�@�C���E�I�[�v���E�t���O
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'

    Fnc_ADO_Open = False

    On Error GoTo Err_Fnc_ADO_Open

    Set Obj_Cn = CurrentProject.Connection '���݂̃f�[�^�x�[�X�֐ڑ�
    Set Obj_RS = New ADODB.Recordset    'ADO���R�[�h�Z�b�g�̃C���X�^���X�쐬
    Obj_RS.Open Str_SQL, Obj_Cn '���R�[�h���o
    
    Fnc_ADO_Open = True

Exit_Fnc_ADO_Open:

    On Error GoTo 0

    Exit Function

Err_Fnc_ADO_Open:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2019/03/22 Add End

''2019/03/25 Add Start
Public Function Fnc_Uru_Chk(Lng_Year As Long) As Integer

    On Error GoTo Err_Fnc_Uru_Chk

    Fnc_Uru_Chk = False

    ''�@�@�N/4    ����؂��  �A��
    If Lng_Year Mod 4 <> 0 Then
        ''����؂�Ȃ� �[�N�Ŗ���
        GoTo Exit_Fnc_Uru_Chk
    Else
        ''�A�@�N/100  ����؂��  �B��
        If Lng_Year Mod 100 = 0 Then
            ''�B�@�N/400
            If Lng_Year Mod 400 <> 0 Then
                ''����؂�Ȃ� �[�N�Ŗ���
                GoTo Exit_Fnc_Uru_Chk
            Else
                ''����؂��  �[�N
            End If
        Else
            ''����؂�Ȃ� �[�N
        End If
    End If

    Fnc_Uru_Chk = True

Exit_Fnc_Uru_Chk:

    On Error GoTo 0

    Exit Function

Err_Fnc_Uru_Chk:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

''2019/04/09 Add Start
Public Function Fnc_DB_Sync() As Integer

    Fnc_DB_Sync = False

    On Error GoTo Err_Fnc_DB_Sync

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������E�J�n")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�P�^�P�Q�j�E�폜�yQD550_TP50_TM00_SUPPLIER_V�z")
    Ret = Fnc_Query_Exec("QD550_TP50_TM00_SUPPLIER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�Q�^�P�Q�j�E�폜�yQD551_TP51_TM01_PRODUCT_MASTER_V�z")
    Ret = Fnc_Query_Exec("QD551_TP51_TM01_PRODUCT_MASTER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�R�^�P�Q�j�E�폜�yQD552_TM02_MATERIAL_V�z")
    Ret = Fnc_Query_Exec("QD552_TM02_MATERIAL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�S�^�P�Q�j�E�폜�yQD553_TM03_TUBE_MATERIAL_V�z")
    Ret = Fnc_Query_Exec("QD553_TM03_TUBE_MATERIAL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�T�^�P�Q�j�E�폜�yQD554_TM04_PRODUCT_SUB_MASTER_V�z")
    Ret = Fnc_Query_Exec("QD554_TM04_PRODUCT_SUB_MASTER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�U�^�P�Q�j�E�폜�yQD559_TM99_MASTER_ALL_V�z")
    Ret = Fnc_Query_Exec("QD559_TM99_MASTER_ALL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�V�^�P�Q�j�E�ǉ��yQA550_TM00_SUPPLIER_V�z")
    Ret = Fnc_Query_Exec("QA550_TM00_SUPPLIER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�W�^�P�Q�j�E�ǉ��yQA551_TM01_PRODUCT_MASTER_V�z")
    Ret = Fnc_Query_Exec("QA551_TM01_PRODUCT_MASTER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�O�X�^�P�Q�j�E�ǉ��yQA552_TM02_MATERIAL_V�z")
    Ret = Fnc_Query_Exec("QA552_TM02_MATERIAL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�P�O�^�P�Q�j�E�ǉ��yQA553_TM03_TUBE_MATERIAL_V�z")
    Ret = Fnc_Query_Exec("QA553_TM03_TUBE_MATERIAL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�P�P�^�P�Q�j�E�ǉ��yQA554_TM04_PRODUCT_SUB_MASTER_V�z")
    Ret = Fnc_Query_Exec("QA554_TM04_PRODUCT_SUB_MASTER_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������i�P�Q�^�P�Q�j�E�ǉ��yQA559_TM99_MASTER_ALL_V�z")
    Ret = Fnc_Query_Exec("QA559_TM99_MASTER_ALL_V", "M00_Public_Module")

    Ret = Fnc_Sys_Msg_Dsp("�}�X�^���������E�I��")

    Fnc_DB_Sync = True

Exit_Fnc_DB_Sync:

    On Error GoTo 0

    Exit Function
    
Err_Fnc_DB_Sync:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function

Public Function Fnc_Fix_Requery(Obj_Form As Object) As Integer
     
    On Error GoTo Err_Fnc_Fix_Requery

    Fnc_Fix_Requery = False


    Dim rst As Recordset
    Dim varBookMark As Variant

    Set rst = Obj_Form.Recordset
''    varBookMark = rst.Form.Bookmark
    varBookMark = rst.Bookmark
    Obj_Form.Requery
    rst.Bookmark = varBookMark


''    Dim recordNumber As Integer
''    recordNumber = Obj_Form.SelTop
''    DoCmd.GoToRecord acDataForm, Obj_Form, acGoTo, recordNumber
     
''    Dim rst As Recordset
''    Dim varBookMark As Variant
''
''    '�t�H�[���̃��R�[�h�Z�b�g��ϐ��ɃZ�b�g���܂�
''    Set rst = Obj_Form.Recordset
''    '���R�[�h�Z�b�g�̃u�b�N�}�[�N���擾���܂�
''    '���ꂪ�ăN�G���[�O�̃J�����g���R�[�h��\���܂�
''    varBookMark = rst.Bookmark
''    '�t�H�[�����ăN�G���[���܂�
''    Obj_Form.Requery
''    '�J�����g���R�[�h��ۑ�����Ă���u�b�N�}�[�N�ɐݒ肵�܂�
''    rst.Bookmark = varBookMark
     
''    Dim varBm As Variant  '�o���A���g�^�̕ϐ����w��
''    varBm = Obj_Form.Bookmark  '���݃��R�[�h�̏�������
''    Obj_Form.Requery  '�ăN�G���̎��s
''    Obj_Form.Bookmark = varBm  '���̃��R�[�h�ɖ߂�


''    Dim headerHeight As Long
''    Dim curTop As Long
''    Dim curRecNum As Long
''    Dim topRecNum As Long
''    'ID�Ƀt�H�[�J�X���ڂ�
''    Obj_Form.SetFocus
''    Set_Object.SetFocus
''
''    '�J�����g���R�[�h���擾
''    curRecNum = Obj_Form.CurrentRecord
''
''    '�t�H�[���w�b�_�[�s�����擾
''    headerHeight = Int(Obj_Form.Section("�t�H�[���w�b�_�[").Height / Obj_Form.Section("�ڍ�").Height)
''
''    '���݂̃Z�N�V�����̏�[����t�H�[���̏�[�܂ł̋����itwip�j���擾
''    curTop = Obj_Form.CurrentSectionTop
''
''    '���ݐ擪�ɕ\������Ă��郌�R�[�h�ԍ����擾
''    topRecNum = curRecNum - (Int(curTop / Obj_Form.Section("�ڍ�").Height) - headerHeight)
''
''    '�ĕ\��
''    Obj_Form.Requery
''
''    '�\���ʒu�̕���
''    DoCmd.GoToRecord acActiveDataObject, , acLast
''    DoCmd.GoToRecord acActiveDataObject, , acGoTo, topRecNum
''    DoCmd.GoToRecord acActiveDataObject, , acGoTo, curRecNum

    Fnc_Fix_Requery = True

Exit_Fnc_Fix_Requery:

    On Error GoTo 0

    Exit Function

Err_Fnc_Fix_Requery:
    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function
''2019/04/09 Add END

''2019/04/10 Add Start
Public Function Fnc_IsExistFile(Str_File_Name As String) As Boolean

    Fnc_IsExistFile = False

    On Error GoTo Err_Fnc_IsExistFile

    Fnc_IsExistFile = (Len(Dir(Str_File_Name)) > 0)

Exit_Fnc_IsExistFile:

    On Error GoTo 0

    Exit Function

Err_Fnc_IsExistFile:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
 
 End Function
''2019/04/10 Add End

''2019/04/12 Add Start
Public Function Fnc_IsExistDirA(a_sFolder As String) As Boolean
    
    Fnc_IsExistDirA = False

    On Error GoTo Err_Fnc_IsExistDirA
    
    If Dir(a_sFolder, vbDirectory) <> "" Then
        '// �t�H���_�����݂���
        Fnc_IsExistDirA = True
    Else
        '// �t�H���_�����݂��Ȃ�
        ''��������
    End If

Exit_Fnc_IsExistDirA:

    On Error GoTo 0

    Exit Function

Err_Fnc_IsExistDirA:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function
''2019/04/12 Add End
''2019/04/17 Add Start
Public Function Fnc_Data_Type_Chk(Obj_Data As Object) As Long

    Fnc_Data_Type_Chk = VarType(Obj_Data)
    
    On Error GoTo Err_Fnc_Data_Type_Chk

    Select Case VarType(Obj_Data)
        Case vbEmpty    '0 �� (��������)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbNull     '1 Null (�����Ȓl)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbInteger  '2 �����^ (Integer)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbLong     '3 �������^ (Long)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbSingle   '4 �P���x���������_�^ (Single)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbDouble   '5 �{���x���������_�^ (Double)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbCurrency '6 �ʉ݌^ (Currency)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbDate     '7 ���t�^ (Date)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbString   '8 ������^ (String)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbObject   '9 �I�u�W�F�N�g
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbError    '10 �G���[�l
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbBoolean  '11 �u�[���^ (Boolean)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbVariant  '12 �o���A���g�^ (Variant) (�o���A���g�^�z��݂̂Ɏg�p)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbDataObject   '13 �f�[�^ �A�N�Z�X �I�u�W�F�N�g
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbDecimal      '14 10 �i���^
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbByte         '17 �o�C�g�^ (Byte)
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbUserDefinedType  '36 ���[�U�[��`�^���܂ރo���A���g�^
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case vbArray            '
            Fnc_Data_Type_Chk = VarType(Obj_Data)
            DoEvents
        Case Else
            Fnc_Data_Type_Chk = -1
            DoEvents
    End Select

Exit_Fnc_Data_Type_Chk:

    On Error GoTo 0

    Exit Function

Err_Fnc_Data_Type_Chk:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2019/04/17 Add End

''2019/05/14 Add Start
Public Function Fnc_DebugPrintFile_OLD(varData As Variant, Optional Str_File_Name As String = "DebugPrint")

    Dim lngFileNum As Long
    Dim strLogFile As String
    
    Dim vardata_Add As Variant
    Dim Str_Wk_File_Name As String
   
    ''2019/05/24 Add Start
    If Fnc_DBG_Mode_Chk() = True Then
    ''2019/05/24 Add End
   
        Str_Wk_File_Name = Str_File_Name & "_" & Format(Now(), "yyyymmdd")
        vardata_Add = Format(Now(), "yyyy/mm/dd hh:mm:ss") & " ->> " & varData
        
        ''strLogFile = CurrentProject.Path & "\" & "DebugPrint.txt"
        strLogFile = CurrentProject.Path & "\" & Str_Wk_File_Name & ".txt"
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, vardata_Add
        Close #lngFileNum
        
        Debug.Print varData

    ''2019/05/24 Add Start
    End If
    ''2019/05/24 Add End

End Function
''2019/05/14 Add End

''2019/05/24 Add Start
Public Function Fnc_DBG_Mode_Chk() As Integer
    
    On Error GoTo Err_Fnc_DBG_Mode_Chk
    
    Fnc_DBG_Mode_Chk = DFirst("Sys_Dbg", "QS80_02_System_Data")
    
Exit_Fnc_DBG_Mode_Chk:
    
    On Error GoTo 0
    
    Exit Function
    
Err_Fnc_DBG_Mode_Chk:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select

End Function
''2019/05/24 Add End

''2019/07/03 Add Start
''Public Function Fnc_DebugPrintFile(Str_PG As String, varData As Variant, Optional Str_File_Name As String = "DebugPrint")
Public Function Fnc_DebugPrintFile(varData As Variant, Str_PG As String, Optional Str_File_Name As String = "DebugPrint")

    Dim lngFileNum As Long
    Dim strLogFile As String
    
    Dim vardata_Add As Variant
    Dim Str_Wk_File_Name As String
   
    If Fnc_DBG_Mode_Chk() = True Then
   
        Str_Wk_File_Name = Str_File_Name & "_" & Format(Now(), "yyyymmdd") & "_2"
        vardata_Add = Format(Now(), "yyyy/mm/dd hh:mm:ss") & "[PG:" & Str_PG & "] ->> " & varData
        
        ''strLogFile = CurrentProject.Path & "\" & "DebugPrint.txt"
        strLogFile = CurrentProject.Path & "\" & Str_Wk_File_Name & ".txt"
        lngFileNum = FreeFile()
        Open strLogFile For Append As #lngFileNum
        Print #lngFileNum, vardata_Add
        Close #lngFileNum
        
        Debug.Print varData

    End If

End Function
''2019/07/03 Add End

''2019/07/09 Add Start
Public Function Fnc_Get_YMD_From2(Str_Year As String, Str_Month As String, Dte_From As Date, Dte_To As Date) As Integer
    
    Dim Int_Year As Integer
    Dim Int_Month As Integer
    Dim Dte_In As Date
    Dim Dte_B_YM As Date
    Dim Dte_N_YM As Date

    Fnc_Get_YMD_From2 = False

    On Error GoTo Err_Fnc_Get_YMD_From2

    ''�N�擾�i�f�t�H���g�E�Z�b�g�܁j
    If Str_Year = "" Then
        Int_Year = Year(Now())
    Else
        Int_Year = Val(Str_Year)
    End If

    ''���擾�i�f�t�H���g�E�Z�b�g�܁j
    If Str_Month = "" Then
        Int_Month = Month(Now())
    Else
        Int_Month = Val(Str_Month)
    End If

    ''�v�Z�ΏہE�N�����Z�b�g
    Dte_In = DateSerial(Int_Year, Int_Month, 1)

    ''�Ώ۔N���̊J�n���v�Z
    Dte_B_YM = DateAdd("m", -1, Dte_In) + 25
    
    ''�Ώ۔N���̏I�����v�Z
    Dte_N_YM = Dte_In + 24

    Dte_From = Dte_B_YM
    Dte_To = Dte_N_YM

    Fnc_Get_YMD_From2 = True

Exit_Fnc_Get_YMD_From2:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_YMD_From2:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function

Public Function Fnc_Get_YMD_From2_From(Str_Year As String, Str_Month As String) As Date

    Dim Dte_From As Date
    Dim Dte_To As Date

    On Error GoTo Err_Fnc_Get_YMD_From2_From


    Ret = Fnc_Get_YMD_From2(Str_Year, Str_Month, Dte_From, Dte_To)

    Fnc_Get_YMD_From2_From = Dte_From


Exit_Fnc_Get_YMD_From2_From:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_YMD_From2_From:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function

Public Function Fnc_Get_YMD_From2_To(Str_Year As String, Str_Month As String) As Date

    Dim Dte_From As Date
    Dim Dte_To As Date

    On Error GoTo Err_Fnc_Get_YMD_From2_To


    Ret = Fnc_Get_YMD_From2(Str_Year, Str_Month, Dte_From, Dte_To)

    Fnc_Get_YMD_From2_To = Dte_To


Exit_Fnc_Get_YMD_From2_To:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_YMD_From2_To:

    Select Case Err
        ''�J�����g�E���R�[�h��
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If
            Resume
            Resume Next
    End Select
End Function
''2019/07/09 Add End


''2019/07/18 Add Start

Public Function Fnc_Get_Calc_Before_YM(ByVal Int_Year As Integer, ByVal Int_Month As Integer) As Date
''    Dim Int_Year As Integer
''    Dim Int_Month As Integer
    Dim Dte_Wk As Date

    On Error GoTo Err_Fnc_Get_Calc_Before_YM

    DoEvents
    
    If Int_Year = 0 Then
        Int_Year = Fnc_Get_This_Year()
    End If

    If Int_Month = 0 Then
        Int_Month = Fnc_Get_This_Month()
    End If

    Dte_Wk = DateSerial(Int_Year, Int_Month, 1)
    
    Dte_Wk = DateAdd("m", -1, Dte_Wk)
    
    DoEvents

    Fnc_Get_Calc_Before_YM = Dte_Wk

Exit_Fnc_Get_Calc_Before_YM:

    On Error GoTo 0

    Exit Function

Err_Fnc_Get_Calc_Before_YM:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function

''2019/07/18 Add End

''2019/10/07 Add Start
Public Function Fnc_Debug_Print_Query_Out(Str_Query_Name As String, Str_PG As String, Optional Str_File_Name As String = "DebugPrint_Query") As Integer

    Dim Obj_DB As Database
    Dim Obj_RS As Recordset
    Dim Int_Row_Cnt As Integer
    Dim Int_Col_Cnt As Integer
    Dim Str_Out_Data    As String

    '�d����e�[�u�����J��
    Set Obj_DB = CurrentDb
    Set Obj_RS = Obj_DB.OpenRecordset(Str_Query_Name)

    Int_Row_Cnt = 0

    If Obj_RS.EOF = False Then

        Ret = Fnc_DebugPrintFile(" �� " & Str_PG & " Data Out Start �� ", Str_PG)

        Int_Row_Cnt = 0
        Str_Out_Data = ""

        Str_Out_Data = ""
        Str_Out_Data = Str_Out_Data & ",""No."" ,"

        For Int_Col_Cnt = 1 To Obj_RS.Fields.Count
            Str_Out_Data = Str_Out_Data & """" & Obj_RS.Fields(Int_Col_Cnt - 1).Name & """ ,"
        Next Int_Col_Cnt

        ''�]���ȕ����폜
        Str_Out_Data = Left(Str_Out_Data, Len(Str_Out_Data) - 1)

        ''���R�[�h�I���ʒu�Ɂy!!�z�ǉ�
        Str_Out_Data = Str_Out_Data & Str_Out_Data & "!!"

        Ret = Fnc_DebugPrintFile(Str_Out_Data, Str_PG)

        '�e���R�[�h���o��
        Do Until Obj_RS.EOF
            Int_Row_Cnt = Int_Row_Cnt + 1

            Str_Out_Data = ""
            Str_Out_Data = Str_Out_Data & ",""" & Int_Row_Cnt & """ ,"

            For Int_Col_Cnt = 1 To Obj_RS.Fields.Count
                Str_Out_Data = Str_Out_Data & """" & Obj_RS.Fields(Int_Col_Cnt - 1) & """ ,"
            Next Int_Col_Cnt

            ''�]���ȕ����폜
            Str_Out_Data = Left(Str_Out_Data, Len(Str_Out_Data) - 1)

            ''���R�[�h�I���ʒu�Ɂy!!�z�ǉ�
            Str_Out_Data = Str_Out_Data & Str_Out_Data & "!!"

            Ret = Fnc_DebugPrintFile(Str_Out_Data, Str_PG)

            Obj_RS.MoveNext
        Loop

        Obj_RS.Close

        Ret = Fnc_DebugPrintFile(" �� " & Str_PG & " Data Out End �� ", Str_PG)

    End If

End Function
''2019/10/07 Add End

''2020/01/31 Add Start
Private Sub Sub_DB_List()
    '�T�v�F�I�u�W�F�N�g�i�e�[�u���A�N�G���[�A�t�H�[���A���|�[�g�A�}�N���A���W���[���j�ꗗ��CSV�o�͂���B

    Dim Obj_DB As AccessObject, dbs As Object
    Dim Lng_Dsn As Long
    Dim strPath As String, strFile As String, strCSVFile As String

    strPath = Application.CurrentProject.Path
    strFile = Application.CurrentProject.Name
    strCSVFile = Left(strFile, InStrRev(strFile, ".")) & "csv"

    Lng_Dsn = FreeFile
    Open strPath & "\" & strCSVFile For Output As #Lng_Dsn

    Set dbs = Application.CurrentData

    'AllTables �R���N�V�������猟��
    For Each Obj_DB In dbs.AllTables
        If Left(Obj_DB.Name, 4) <> "MSys" Then
            Write #Lng_Dsn, strPath, strFile, "�e�[�u��", Obj_DB.Name
        End If
    Next Obj_DB

    'AllQueries �R���N�V�������猟��
    For Each Obj_DB In dbs.AllQueries
        Write #Lng_Dsn, strPath, strFile, "�N�G���[", Obj_DB.Name
    Next Obj_DB
    Set dbs = Nothing

    '
    Set dbs = Application.CurrentProject

    'AllForms �R���N�V�������猟��
    For Each Obj_DB In dbs.AllForms
        Write #Lng_Dsn, strPath, strFile, "�t�H�[��", Obj_DB.Name
    Next Obj_DB

    'AllReports �R���N�V�������猟��
    For Each Obj_DB In dbs.AllReports
        Write #Lng_Dsn, strPath, strFile, "���|�[�g", Obj_DB.Name
    Next Obj_DB

    'AllMacros �R���N�V�������猟��
    For Each Obj_DB In dbs.AllMacros
        Write #Lng_Dsn, strPath, strFile, "�}�N��", Obj_DB.Name
    Next Obj_DB

    'AllModules �R���N�V�������猟��
    For Each Obj_DB In dbs.AllModules
        Write #Lng_Dsn, strPath, strFile, "���W���[��", Obj_DB.Name
    Next Obj_DB

    Set dbs = Nothing
    Close #Lng_Dsn

''End Sub
End Sub

''2020/01/31 Add End

''2020/02/18 Add Start
Public Function Fnc_QD_F03_00_03_06_TD04_Material_Plan() As Integer
'********************************************************************************
'*
'*  �yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j
'*
'*-------------------------------------------------------------------------------
'*
'*   ����
'*
'*-------------------------------------------------------------------------------
'*
'*   �߂�l
'*           True       :   ����I��
'*
'*           False      :   �X�V��
'*
'********************************************************************************
'
''�Q�lSQL
''DELETE TD04_Material_Plan.*
''FROM TD04_Material_Plan;

    Dim Obj_RS As New ADODB.Recordset
    Dim Lng_D_Cnt As Long
    Dim Lng_P_Cnt As Long
    
    Dim Str_Msg As String
    
    On Error GoTo Err_Fnc_QD_F03_00_03_06_TD04_Material_Plan
    
    Fnc_QD_F03_00_03_06_TD04_Material_Plan = False

    Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�J�n"
    Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
    Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)

    Lng_P_Cnt = 0

    ''DB�I�[�v���iTD04_Material_Plan�j
    Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�EDB�I�[�v���iTD04_Material_Plan�j"
    Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
    Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
    Obj_RS.Open "TD04_Material_Plan", CurrentProject.Connection

    ''�f�[�^�L���m�F
    Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�f�[�^�L���m�F"
    Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
    Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
    If Obj_RS.EOF = False Then
        Obj_RS.MoveLast

        ''�f�[�^�����擾
        Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�f�[�^�����擾"
        Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
        Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
        Lng_D_Cnt = Obj_RS.RecordCount
    
        Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�f�[�^����:" & Lng_D_Cnt & "��"
        Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
        Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)

        Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�S�Ė����Ȃ�܂Ń��[�v"
        Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
        Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
        Do While Obj_RS.EOF = False
            Lng_P_Cnt = Lng_P_Cnt + 1
    
            Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�擪�擾"
            Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
            ''Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
            Obj_RS.MoveFirst

            Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E���R�[�h�폜(" & Lng_P_Cnt & "/" & Lng_D_Cnt & ")"
            Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
            Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
            Obj_RS.Delete
        Loop
        Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�S�Ė����Ȃ���"
        Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
        Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
    
    End If

    Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�EDB�N���[�Y�iTD04_Material_Plan�j"
    Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
    Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)
    Obj_RS.Close
    Set Obj_RS = Nothing

    Fnc_QD_F03_00_03_06_TD04_Material_Plan = True

Exit_Fnc_QD_F03_00_03_06_TD04_Material_Plan:
    
    Str_Msg = "�yTD04_Material_Plan�z�e�[�u���E�f�[�^�S�폜�iAccess�o�O�Ή��j�E�I��"
    Ret2 = Fnc_DebugPrintFile(Str_Msg, "M00_Public_Module:Fnc_QD_F03_00_03_06_TD04_Material_Plan")
    Ret2 = Fnc_Sys_Msg_Dsp(Str_Msg)

    On Error GoTo 0

    Exit Function

Err_Fnc_QD_F03_00_03_06_TD04_Material_Plan:

    Select Case Err
        Case Else                                                               '��L�ȊO�̃G���[
            If DBEngine.Errors.Count > 0 Then
                ' Errors �R���N�V������񋓂��܂��B
                For Each Errloop In DBEngine.Errors
                    MsgBox "Error number:" & Errloop.Number & _
                        vbCr & Errloop.Description
                Next Errloop
            End If

            Resume Next
    End Select

End Function
''2020/02/18 Add End
