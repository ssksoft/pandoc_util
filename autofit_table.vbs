dim fso
set fso = createObject("Scripting.FileSystemObject")

Const END_OF_STORY = 6
Const wdPageBreak = 7
 
Set objWord = CreateObject("Word.Application")
 
objWord.Visible = True
 
'ワードドキュメントを新規作成
filename = fso.getParentFolderName(WScript.ScriptFullName) & "\sample.docx"
Set objDoc = objWord.Documents.Open(filename)
Set objSelection = objWord.Selection
 
'テーブルの初期サイズの指定
objDoc.Tables(1).Rows.Alignment = 1