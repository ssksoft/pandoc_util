Const END_OF_STORY = 6
Const wdPageBreak = 7
 
Set objWord = CreateObject("Word.Application")
 
objWord.Visible = True
 
'ワードドキュメントを新規作成
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
 
'ページ番号をフッターに書き込む処理
Set objSection = objDoc.Sections(1)
Set objFooters = objSection.Footers(1).PageNumbers
objFooters.Add(1)
 
'フォントの指定
objSelection.Font.Name = "ＭＳ 明朝"
 
'フォントサイズ
objSelection.Font.Size = "12"
 
'本文の入力
objSelection.TypeText "Win32_Serviceサービスの一覧表示"
 
'文字入力後にEnterキーを押すのと同じ意味
objSelection.TypeParagraph()
 
'フォントサイズ
objSelection.Font.Size = "10.5"
 
'日付を本文に書き込みます
objSelection.TypeText "" & Date()
 
objSelection.TypeParagraph()
 
'テーブルを作成する
Set objRange = objSelection.Range
 
'テーブルの初期サイズの指定
objDoc.Tables.Add objRange,1,3
Set objTable = objDoc.Tables(1)