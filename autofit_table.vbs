Const END_OF_STORY = 6
Const wdPageBreak = 7
 
Set objWord = CreateObject("Word.Application")
 
objWord.Visible = True
 
'���[�h�h�L�������g��V�K�쐬
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
 
'�y�[�W�ԍ����t�b�^�[�ɏ������ޏ���
Set objSection = objDoc.Sections(1)
Set objFooters = objSection.Footers(1).PageNumbers
objFooters.Add(1)
 
'�t�H���g�̎w��
objSelection.Font.Name = "�l�r ����"
 
'�t�H���g�T�C�Y
objSelection.Font.Size = "12"
 
'�{���̓���
objSelection.TypeText "Win32_Service�T�[�r�X�̈ꗗ�\��"
 
'�������͌��Enter�L�[�������̂Ɠ����Ӗ�
objSelection.TypeParagraph()
 
'�t�H���g�T�C�Y
objSelection.Font.Size = "10.5"
 
'���t��{���ɏ������݂܂�
objSelection.TypeText "" & Date()
 
objSelection.TypeParagraph()
 
'�e�[�u�����쐬����
Set objRange = objSelection.Range
 
'�e�[�u���̏����T�C�Y�̎w��
objDoc.Tables.Add objRange,1,3
Set objTable = objDoc.Tables(1)