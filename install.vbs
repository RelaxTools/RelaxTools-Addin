' -------------------------------------------------------------------------------
' RelaxTools-Addin �C���X�g�[���X�N���v�g Ver.1.0.6
' -------------------------------------------------------------------------------
' �Q�l�T�C�g
' ���� SE �̂Ԃ₫
' VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@
' http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html
' �C��
'   1.0.6 �C���X�g�[���p�X�� Application.UserLibraryPath �𗘗p����悤�ɏC���B
'   1.0.5 �����u�b�N���Q�Ɨp�ɊJ��VBS���C���X�g�[������悤�C���B
'   1.0.4 �}���`�v���Z�X�pVBS���s�v�ɂȂ����̂ō폜�B
'   1.0.3 �}���`�v���Z�X�pVBS���R�s�[����悤�C���B
'   1.0.3 images �t�H���_���R�s�[����悤�ɏC���B
'   1.0.2 Windows Update �ɂ� �C���^�[�l�b�g���擾�����A�h�C���t�@�C���� Excel �ɂēǂݍ��܂�Ȃ��ꍇ�ɑΉ��B
'         �x���ƃv���p�e�B�E�B���h�E��\�����āu�u���b�N�����v�����肢����悤�ɂ����B
' -------------------------------------------------------------------------------
Option Explicit
On Error Resume Next

Dim installPath 
Dim addInName 
Dim addInFileName 
Dim objExcel 
Dim objAddin
Dim imageFolder
Dim appFile
Dim objWshShell
Dim objFileSys
Dim strPath
Dim objFolder
Dim objFile

'�A�h�C������ݒ� 
addInName = "RelaxTools Addin" 
addInFileName = "Relaxtools.xlam"
appFile = "rlxAliasOpen.vbs"

Set objWshShell = CreateObject("WScript.Shell") 
Set objFileSys = CreateObject("Scripting.FileSystemObject")

IF Not objFileSys.FileExists(addInFileName) THEN
    MsgBox "Zip�t�@�C����W�J���Ă�����s���Ă��������B", vbExclamation, addInName 
    WScript.Quit 
END IF

IF MsgBox(addInName & " ���C���X�g�[�����܂����H" & vbCrLf &  "Version 4.0.0 �ȍ~�Ƃ���ȑO�ł͐ݒ肪�����p����܂���̂ł��������������B", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

'Excel �C���X�^���X�� 
With CreateObject("Excel.Application") 

    '�C���X�g�[����p�X�̍쐬 
    strPath = .UserLibraryPath
    imageFolder = objWshShell.SpecialFolders("Appdata") & "\RelaxTools-Addin\"

    '�C���X�g�[���t�H���_���Ȃ��ꍇ�͍쐬
    IF Not objFileSys.FolderExists(strPath) THEN
        objFileSys.CreateFolder(strPath)
    END IF

    installPath = strPath & addInFileName

    '�t�@�C���R�s�[(�㏑��) 
    objFileSys.CopyFile  addInFileName ,installPath , True

    '�C���[�W�t�H���_���Ȃ��ꍇ�͍쐬
    IF Not objFileSys.FolderExists(imageFolder) THEN
        objFileSys.CreateFolder(imageFolder)
    END IF

    '�C���[�W�t�H���_���R�s�[(�㏑��) 
    objFileSys.CopyFolder  "Source\customUI\images" ,imageFolder , True

    '�t�@�C�����R�s�[(�㏑��) 
    objFileSys.CopyFile  appFile, imageFolder & appFile, True

    '�A�h�C���o�^ 
    .Workbooks.Add
    Set objAddin = .AddIns.Add(installPath, True) 
    objAddin.Installed = True

    'Excel �I�� 
    .Quit

End WIth

IF Err.Number = 0 THEN 
    MsgBox "�A�h�C���̃C���X�g�[�����I�����܂����B", vbInformation, addInName 

    '�v���p�e�B�t�@�C���\��
    CreateObject("Shell.Application").NameSpace(strPath).ParseName(addInFileName).InvokeVerb("properties")
    MsgBox "�C���^�[�l�b�g����擾�����t�@�C����Excel���u���b�N�����ꍇ������܂��B" & vbCrlf & "�v���p�e�B�E�B���h�E���J���܂��̂Łu�u���b�N�̉����v���s���Ă��������B" & vbCrLf & vbCrLf & "�v���p�e�B�Ɂu�u���b�N�̉����v���\������Ȃ��ꍇ�͓��ɑ���̕K�v�͂���܂���B", vbExclamation, addInName 

ELSE 
    MsgBox "�G���[���������܂����B" & vbCrLF & "Excel���N�����Ă���ꍇ�͏I�����Ă��������B", vbExclamation, addInName 
    WScript.Quit 
End IF

If MsgBox("�G�N�X�v���[���E�N���b�N(�����u�b�N���Q�Ɨp�ɊJ��)��L���ɂ��܂����H" & vbCrLf & "���s�ɂ͊Ǘ��Ҍ������K�v�ł��B", vbYesNo + vbQuestion, addInName) <> vbNo Then 
    objWshShell.Run "rlxAliasOpen.vbs /install", 1, true
End IF

If MsgBox("�G�N�X�v���[���E�N���b�N(Excel�̓ǂݎ���p)��L���ɂ��܂����H" & vbCrLf & "���s�ɂ͊Ǘ��Ҍ������K�v�ł��B", vbYesNo + vbQuestion, addInName) = vbNo Then 
    WScript.Quit 
End IF

objWshShell.Run "ExcelReadOnly.vbs", 1, true

Set objFileSys = Nothing
Set objWshShell = Nothing 
