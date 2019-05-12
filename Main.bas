Attribute VB_Name = "Main"
Dim jColumn As Long     '���w�̐g����
Dim jCell As range      '���w�̐g���Z��
Dim hCell As range      '���Z�̖��O�Z��

Sub Main()
    '---�ϐ��Q---
    Dim hname As Variant        '���Z���̖��O
    Dim resultSheet As String   '���ʏo�̓V�[�g�̖��O
    
    Dim hlist() As Variant      '���Z����̔z��
    
    Dim row As Integer          '�o�͍s���J�E���^
    
    '-----�������珈���J�n-----
    '�o�͂���V�[�g���̊i�[
    resultSheet = "�o�͌���"
    
    '�t�H�[������������
    Form.Show
    
    '���w�V�[�g���A�N�e�B�u�ɂ��A�g���擪�Z���A��C���f�b�N�X���擾����
    Worksheets(Form.jSheet.Text).Activate
    On Error GoTo SelectError
    Set jCell = SelectCell("���w�V�[�g�̐g���擪�Z���i���ږ��������j��I�����Ă��������B")
    jColumn = jCell.Column
    
    '���Z�V�[�g���A�N�e�B�u�ɂ��A����擪�Z�����擾����
    Worksheets(Form.hSheet.Text).Activate
    On Error GoTo SelectError
    Set hCell = SelectCell("���Z����̎����擪�Z���i���ږ��������j��I�����Ă��������B")
    
    '��ʍX�V��~
    Application.ScreenUpdating = False
    
    '���Z����̔z����擾����
    hlist() = gethNames()
    
    '���ʏo�̓V�[�g���A�N�e�B�u�ɂ���
    On Error GoTo Sheetrack
    
    '�o�͐擪�s���i�[����
    row = 2
    
    '�e���Z���k���Ō������\�b�h���Ăяo���A�߂�l�����������̂ɂ��ăV�[�g�ɏo�͂���
    For Each hname In hlist
        Dim std As Student
        '�������\�b�h�Ăяo��
        Set std = getStudent(hname)
        
        If Not std Is Nothing Then
            '���ڂɖ��O���i�[
            Worksheets(resultSheet).Cells(row, 1).Value = std.name
            '���ڈȍ~�ɐg���̏d���i�[
            range(Worksheets(resultSheet).Cells(row, 2), Worksheets(resultSheet).Cells(row + 1, 10)).Value = std.getData
            '�s�����X�V
            row = row + 2
        End If
    Next
    '�o�͌��ʃV�[�g���A�N�e�B�u�ɂ���
    Worksheets(resultSheet).Activate
    
    '�t�H�[�������
    Unload Form
    
    MsgBox "�������������܂����B"
    Exit Sub
'�Z�����I������Ȃ������ꍇ�A�������I������
SelectError:
    MsgBox "�Z���̎擾�Ɏ��s���܂����B"
    End
'�o�̓V�[�g���Ȃ��ꍇ�A�쐬����
Sheetrack:
    Worksheets.Add.name = resultSheet
    Worksheets(resultSheet).range("A1:J1").Value = Array("����", "��1", "��2", "��3", "��4", "��5", "��6", "��1", "��2", "��3")
    Resume
End Sub
'�Z����I�������A���̍���̃Z����ԋp����
Function SelectCell(ByVal txt As String) As range
    Dim rng As range
    Set rng = Application.InputBox(txt, "�Z���I��", Type:=8)
    Set SelectCell = rng.Cells(1)
End Function
'���Z����̑S���O��z��Ŏ擾����
Function gethNames() As Variant()
    Dim nameCells As range
    Dim arr() As Variant
    
    Set nameCells = range(hCell, hCell.End(xlDown))
    arr() = nameCells.Value
    gethNames = arr()
End Function
'���w�V�[�g�̌������s���A���O�������Student�I�u�W�F�N�g�𐶐��A�Ȃ����Nothing���i�[
Function getStudent(name As Variant) As Student
    Dim rng As range
    
    '���w�V�[�g�������Ō�������
    Set rng = Worksheets(Form.jSheet.Text).Cells.Find(name, Cells(1, 1), xlValues, xlWhole, xlByColumns, xlNext, False, False, False)
    
    '�������ʂ��������ꍇ�AStudent�I�u�W�F�N�g�𐶐�����
    If Not rng Is Nothing Then
        Dim i As Integer    '�J�E���^
        Dim row As Long     '�s�l
        Dim dataCell As range   '�g���̏d�̐擪�Z��
        Dim person As Student   '�ԋp����I�u�W�F�N�g
        Dim arr(17) As Variant  '�g���̏d�z��
        Dim datas() As Variant  '�f�[�^�擾�p�z��
        
        '�f�[�^�̍s�l���擾
        row = rng.row
        
        '�g���̏d�̐擪�Z�����擾
        Set dataCell = Worksheets(Form.jSheet.Text).Cells(row, jColumn)
        
        '�͈͂�3�܂Ŋg�債�A�l���擾
        datas = dataCell.Resize(1, 18).Value
        
        '�l��1�����z��ɕϊ�����
        For i = 0 To 17
            arr(i) = datas(1, i + 1)
        Next
        
        'Student�I�u�W�F�N�g�𐶐����āA���O�Ɛg���̏d���i�[����
        Set person = New Student
        person.name = rng.Value
        Call person.setData(arr)
        
        '�I�u�W�F�N�g�̕ԋp
        Set getStudent = person
    Else
        '�������ʂ��Ȃ��ꍇ�ANothing��ԋp
        Set getStudent = Nothing
    End If
End Function
