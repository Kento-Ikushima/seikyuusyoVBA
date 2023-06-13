Attribute VB_Name = "Module1"
Option Explicit

Sub �������쐬()
    Const HAN_HIDUKE_CLM As Long = 1 '�u�̔��v���[�N�V�[�g�́u���t�v�̗�
    Const HAN_KOKYAKU_CLM As Long = 2 '�u�̔��v���[�N�V�[�g�́u�ڋq�v�̗�
    Const HAN_SYOHIN_CLM As Long = 3 '�u�̔��v���[�N�V�[�g�́u���i�v�̗�
    Const HAN_TANKA_CLM As Long = 4 '�u�̔��v���[�N�V�[�g�́u�P���v�̗�
    Const HAN_SURYO_CLM As Long = 5 '�u�̔��v���[�N�V�[�g�́u���ʁv�̗�
    Const HAN_KINGAKU_CLM As Long = 6 '�u�̔��v���[�N�V�[�g�́u���z�v�̗�
    
    Const SEI_HIDUKE_CLM As Long = 1 '�������̃��[�N�V�[�g�́u���t�v�̗�
    Const SEI_SYOHIN_CLM As Long = 2 '�������̃��[�N�V�[�g�́u���i�v�̗�
    Const SEI_TANKA_CLM As Long = 3 '�������̃��[�N�V�[�g�́u�P���v�̗�
    Const SEI_SURYO_CLM As Long = 4 '�������̃��[�N�V�[�g�́u���ʁv�̗�
    Const SEI_KINGAKU_CLM As Long = 5 '�������̃��[�N�V�[�g�́u���z�v�̗�
    
    Const SEITP_WSNM As String = "���������`" '�������e���v���[�g�̃��[�N�V�[�g��
    Const ATESAKI_ADRS As String = "A6" '�������̈���̃Z���Ԓn
    Const HAKKOBI_ADRS As String = "E2" '�������̔��s���̃Z���Ԓn
    

    Dim i As Long '�u�̔��v���[�N�V�[�g�̕\�̏����p�J�E���g�ϐ�
    Dim Cnt As Long '�������̃��[�N�V�[�g�̕\�̏����p�ϐ�
    Dim Kokyaku As String '���������쐬����ڋq��
    Dim HanKiten As Range '�u�̔��v���[�N�V�[�g�̕\�̊�_�Z��
    Dim SeiKiten As Range '�������̃��[�N�V�[�g�̕\�̊�_�Z��
    Dim sheetExists As Boolean '�����̃V�[�g�����݂��邩�̃t���O
    Dim ws As Object '���[�N�V�[�g�𑖍����邽�߂̕ϐ�
    
    
    Cnt = 1 '�������̃��[�N�V�[�g�̕\�̐擪�s�̒l�ɏ�����
    Kokyaku = myForm.myComboBox.Value '�t�H�[���̃h���b�v�_�E���őI�񂾌ڋq��ݒ�
    sheetExists = False '���݂��Ȃ��A�ɏ�����
  
    '�����̃V�[�g�����݂��邩�m�F
    For Each ws In Worksheets
        If ws.Name = Kokyaku Then
            sheetExists = True
            Exit For
        End If
    Next ws

    '�����̃V�[�g�����݂���ꍇ�A���b�Z�[�W��\�����ďI��
    If sheetExists Then
        MsgBox "���̐������͂��łɔ��s�ς݂ł��B"
        Exit Sub
    End If
    
    
    '���[�N�V�[�g�u���������`�v�𖖔��ɃR�s�[
    Worksheets("���������`").Copy After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = Kokyaku      '���[�N�V�[�g���ݒ�
    Worksheets(Kokyaku).Range(ATESAKI_ADRS).Value = Kokyaku  '����̐ݒ�
    Worksheets(Kokyaku).Range(HAKKOBI_ADRS).Value = Date     '���s���̓���
    
    Set HanKiten = Worksheets("�̔�").Range("A4") '�u�̔��v���[�N�V�[�g�̕\�̊�_�Z����ݒ�
    Set SeiKiten = Worksheets(Kokyaku).Range("A12") '�������̃��[�N�V�[�g�̕\�̊�_�Z����ݒ�
    
    '�w�肵���̔��f�[�^�𐿋����փR�s�[
    For i = 1 To HanKiten.CurrentRegion.Rows.Count - 1
      If HanKiten.Cells(i, HAN_KOKYAKU_CLM).Value = Kokyaku Then
        SeiKiten.Cells(Cnt, SEI_HIDUKE_CLM).Value = HanKiten.Cells(i, HAN_HIDUKE_CLM).Value '���t
        SeiKiten.Cells(Cnt, SEI_SYOHIN_CLM).Value = HanKiten.Cells(i, HAN_SYOHIN_CLM).Value '���i
        SeiKiten.Cells(Cnt, SEI_TANKA_CLM).Value = HanKiten.Cells(i, HAN_TANKA_CLM).Value '�P��
        SeiKiten.Cells(Cnt, SEI_SURYO_CLM).Value = HanKiten.Cells(i, HAN_SURYO_CLM).Value '����
        SeiKiten.Cells(Cnt, SEI_KINGAKU_CLM).Value = HanKiten.Cells(, HAN_KINGAKU_CLM).Value '���z
        
        Cnt = Cnt + 1 '�������̃��[�N�V�[�g�̕\�̃R�s�[���1�i�߂�
      End If
    Next
    
    '�t�H�[�����A�����[�h����
    Unload myForm

    'PDF�쐬�̊m�F���b�Z�[�W��\��
    Dim response As Integer
    response = MsgBox("������PDF���쐬���܂����H", vbQuestion + vbYesNo, "�m�F")
    
    '���[�U�[�̑I���ɉ����ď��������s
    If response = vbYes Then
        'PDF�쐬���������s����֐����Ăяo��
        CreatePDF Kokyaku
    End If
End Sub

Sub CreatePDF(ByVal Kokyaku As String)
    ' ���������[�N�V�[�g�͈̔͂�I��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(Kokyaku)
    Dim lastRow As Long
    Const SEI_HIDUKE_CLM As Long = 1 '�������̃��[�N�V�[�g�́u���t�v�̗�
    Const SEI_KINGAKU_CLM As Long = 5 '�������̃��[�N�V�[�g�́u���z�v�̗�
    lastRow = ws.Cells(ws.Rows.Count, SEI_HIDUKE_CLM).End(xlUp).Row
    Dim printRange As Range
    Set printRange = ws.Range(ws.Cells(SEI_HIDUKE_CLM, 1), ws.Cells(lastRow, SEI_KINGAKU_CLM))
    
    ' PDF�t�@�C����ۑ�����p�X���w��
    Dim tempFilePath As String
    tempFilePath = "C:\�l\����\install\VBA�׋��p\sample\TempPDF_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"

    '�uMicrosoft Print to PDF�v�v�����^�[��ݒ�
    Application.ActivePrinter = "Microsoft Print to PDF on Ne01:"
    printRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempFilePath, Quality:=xlQualityStandard
    
    ' PDF�쐬��A���b�Z�[�W��\��
    MsgBox "PDF�쐬���������܂����B"
End Sub

Sub �t�H�[���p��() '�{�^���ɖ��ߍ��ޗp
    myForm.Show
End Sub
Sub �������폜()
    ' �m�F���b�Z�[�W��\�����A�폜��I�������ꍇ�̂ݍ폜���������s����
    Dim response As Integer
    response = MsgBox("�쐬�������������ꊇ�ō폜���܂����H", vbQuestion + vbYesNo, "�m�F")
    
    If response = vbYes Then
        Dim ws As Worksheet
        Application.DisplayAlerts = False ' �m�F���b�Z�[�W���\���ɂ���
        
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> "�̔�" And ws.Name <> "���������`" And ws.Name <> "�ݒ�" Then
                ws.Delete
            End If
        Next ws
        
        Application.DisplayAlerts = True ' �m�F���b�Z�[�W���ĕ\������
        
        MsgBox "�폜���������܂����B"
    End If
End Sub
'���������������Ă��邩�̎���




