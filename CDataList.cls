VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------------
'データベーステーブルのような実データを保持するクラスです。
'外部からのアクセスは1次元配列のように扱えますが、
'内部的には2重の連想配列となっていて、1段目の連想配列は
'Primaryに指定した項目の値をキー値として、値がそのキーが存在する
'データの連想配列となります。
'2段目の連想配列は、FieldNamesに指定した各フィールド名を
'キー値として､それに対応する値となります｡

'Primaryに指定する項目は必ずUniqueな値である必要があります。

'このデータクラスを使うためには以下の手順が必要です。
'1.Fields.addをコールし、データの各項目名を追加します。
'2.主キーとなる項目をFieldNamesから選んでPrimaryに設定します。
'3.OpenDataプロシージャをコールしてデータ入力可能状態にします。

'※Primaryに設定できるのはFldNamesにあるものだけです。
'   よって各項目名の追加をする前にPrimaryの設定はできません。
'※OpenDataされていない場合、データ入力はすべて却下されます。

'Copyright (c) 2016 TskGunGun
'Released under the MIT license
'https://github.com/TSKGunGun/VBADataList
'
'本ソフトウェアは自由な変更を認めます。また、コピー及び再配布・商用利用など自由に使ってください。
'再配布する場合、本ライセンス全文をソースコード、もしくはライセンス表示ファイルなどに掲載してください。
'------------------------------------------------------

'------------------------------------------------
'Class Variables
'------------------------------------------------
Private DataList As Scripting.Dictionary
Private FldNames As Collection
Private Primary As Variant
Private DataOpen As Boolean

'------------------------------------------------
'Class Procedures
'------------------------------------------------
'Constructor
Private Sub Class_initialize()
    Set DataList = New Scripting.Dictionary
    Set FldNames = New Collection
End Sub

'Destcuctor
Private Sub Class_Terminated()
    Set DataList = Nothing
    Set FldNames = Nothing
End Sub

'------------------------------------------------
'Public Procedures
'------------------------------------------------
'プライマリキーの項目名を取得
Public Property Get PrimaryKey() As Variant
    PrimaryKey = Primary
End Property

'プライマリキーとなる項目名を設定
'フィールド名リストになければ設定不可能
'Open済みならプライマリ変更不可能
Public Property Let PrimaryKey(FldName As Variant)
    If Not IsFldExist(FldName) Then Exit Property
    If IsOpen Then Exit Property
    
    Primary = FldName
End Property

'項目名の一覧を保持するコレクションを取得
'メソッドチェインで.addをコールすれば項目を追加できます
Public Property Get Fields() As Collection
    Set Fields = FldNames
End Property

'項目名の一覧を纏めて追加(置き換え)
'メソッドチェインで.addをコールすれば項目を追加できます
Public Property Let Fields(FieldList As Collection)
    Set FldNames = FieldList
End Property

'指定したフィールド名は存在するか？
Public Function IsFldExist(SearchName As Variant) As Boolean
    Dim FldName As Variant
    For Each FldName In FldNames
        If SearchName = FldName Then
            IsFldExist = True
            Exit Function
        End If
    Next
End Function

Public Function IsOpen() As Boolean
    IsOpen = DataOpen
End Function

Public Sub OpenData()
    DataOpen = True
End Sub

'指定アイテムを追加する
'同じPrimaryKeyのアイテムは追加を却下する
'GetBlankItemで取得したScripting.Dictionaryが追加できます
Public Sub Add(item As Scripting.Dictionary)
    If Not IsOpen Then Exit Sub
    
    If DataList.Exists(item.item(Primary)) Then Exit Sub
    
    DataList.Add item.item(Primary), item
End Sub

'同じPrimaryKeyをもつデータが存在すれば更新し、存在しなければ追加を行う
Public Sub Upsert(PrimaryKey As Variant, FieldName As Variant, Value As Variant)
    If Not IsOpen Then Exit Sub
    
    If IsKey(PrimaryKey) Then
        SetValue PrimaryKey, FieldName, Value
    Else
        Dim NewData As Scripting.Dictionary
        Set NewData = AddNew(PrimaryKey)
        NewData(FieldName) = Value
    End If
    
End Sub

'指定したPrimaryKeyを持つ新しいデータを追加し、そのデータのオブジェクトの参照を返す
Public Function AddNew(PrimaryKey As Variant) As Scripting.Dictionary
    If Not IsOpen Then Exit Function
    If DataList.Exists(PrimaryKey) Then Exit Function
        
    Dim Blank As Scripting.Dictionary
    Set Blank = GetBlankData
    Blank.item(Primary) = PrimaryKey
    
    DataList.Add Blank.item(Primary), Blank
    
    Set AddNew = Blank
End Function

'設定した項目名が各キーに設定され,値は空の連想配列を取得する
Public Function GetBlankData() As Scripting.Dictionary
    Dim Blank As Scripting.Dictionary
    Set Blank = New Scripting.Dictionary
    
    Dim FldName As Variant
    For Each FldName In FldNames
        Blank.Add FldName, ""
    Next
    
    Set GetBlankData = Blank
End Function

'このクラスが保持するPrimaryKeyの一覧を取得する
Public Function Keys() As Collection
    Dim Key As Variant
    Dim PrimaryKeys As Collection
    Set PrimaryKeys = New Collection
    For Each Key In DataList.Keys
        PrimaryKeys.Add Key
    Next
    Set Keys = PrimaryKeys
End Function

'指定のPrimaryKeyが存在するかを取得する
Public Function IsKey(PrimaryKey As Variant) As Boolean
    IsKey = DataList.Exists(PrimaryKey)
End Function

'指定のPrimaryKeyを持つ連想配列の参照を返す
Public Function GetAt(PrimaryKey As Variant) As Scripting.Dictionary
    Set GetAt = DataList.item(PrimaryKey)
End Function

'指定フィールド名に指定された値が設定された項目を検索し、最初に見つかったデータのPrimaryKeyを取得する
Public Function Find(FieldName As Variant, Value As Variant) As Variant
    Dim Key As Variant
    For Each Key In Keys
        If GetValue(Key, FieldName) = Value Then
            Find = Key
            Exit Function
        End If
    Next
End Function

'指定フィールド名に指定された値が設定された項目を検索し、見つかったすべてのデータのPrimaryKeyを取得する
Public Function FindAll(FieldName As Variant, Value As Variant) As Collection
    Dim ResKeys As Collection
    Set ResKeys = New Collection
    
    Dim Key As Variant
    For Each Key In Keys
        If Value = GetValue(Key, FieldName) Then
            ResKeys.Add Key
        End If
    Next
    
    Set FindAll = ResKeys
End Function

'指定フィールドに指定された値が設定された項目を抽出する。
Public Function Where(FieldName As Variant, Value As Variant) As CDataList
    Dim ResList As CDataList
    Set ResList = New CDataList
    ResList.Fields = Fields
    ResList.PrimaryKey = PrimaryKey
    ResList.OpenData
    
    Dim Key As Variant
    For Each Key In Keys
        If Value = GetValue(Key, FieldName) Then
           ResList.Add DataList.item(Key)
        End If
    Next
    
    Set Where = ResList
End Function

'指定PrimaryKeyの指定項目の値を取得する
Public Function GetValue(PrimaryKey As Variant, FieldName As Variant) As Variant
    GetValue = GetAt(PrimaryKey).item(FieldName)
End Function

'指定PrimaryKeyの指定項目の値を設定する
Public Sub SetValue(PrimaryKey As Variant, FieldName As Variant, Value As Variant)
    GetAt(PrimaryKey).item(FieldName) = Value
End Sub

'このクラスのデータをADODB.RecordSetに変換して取得する
Public Function GetRecordSet() As ADODB.Recordset
    Set GetRecordSet = RecordSetConv(Keys)
End Function

'このクラスのデータをADODB.RecordSetに変換して取得する(ソート付き)
Public Function GetRecordSet_with_Sort(FieldName As Variant, Optional ASC As Boolean = True) As ADODB.Recordset
    Set GetRecordSet_with_Sort = RecordSetConv(GetSortedKey(FieldName, ASC))
End Function

'データ数を取得する
Public Function Size() As Integer
    Size = DataList.Count
End Function

'指定プライマリキーの指定フィールドの値を１つインクリメントする
'データが空であれば１をセットする
Public Sub Increment(PrimaryKey As Variant, FieldName As Variant)
    If Not IsOpen Then Exit Sub
    If Not IsKey(PrimaryKey) Then AddNew (PrimaryKey)
    
    If GetValue(PrimaryKey, FieldName) = "" Then
        SetValue PrimaryKey, FieldName, 1
    Else
        SetValue PrimaryKey, FieldName, GetValue(PrimaryKey, FieldName) + 1
    End If
End Sub

'すべてのデータを削除する
Public Sub RemoveAll()
    DataList.RemoveAll
End Sub

'指定PrimaryKeyのデータを削除する
Public Sub Remove(PrimaryKey As Variant)
    DataList.Remove (Key)
End Sub

'指定フィールドのユニークな値の集合を返す
'SQLのGROUP BY と同じ
Public Function UniqueValues(FieldName As Variant) As Collection
    Dim Uniques As Collection
    Set Uniques = New Collection
    
    Dim Key As Variant
    For Each Key In Keys
        Dim Value As Variant: Value = GetValue(Key, FieldName)
        
        If Not IsExistOfCollection(Uniques, Value) Then
            Uniques.Add Value
        End If
    Next

    Set UniqueValues = Uniques
End Function

'デフォルトのシートへの出力です。
'1行目がフィールド名、２行目からデータが始まるようなシートになります。
'それ以外の形式で出力した場合は自作してください。
Public Sub WriteSheet(ByRef Sheet As Worksheet)
    Dim RIndex, CIndex As Integer: RIndex = 1: CIndex = 1

    'OutPut FieldNames
    Dim FldName As Variant
    For Each FldName In FldNames
        Sheet.Cells(RIndex, CIndex) = FldName
        CIndex = CIndex + 1
    Next
    
    'OutPut Datas
    RIndex = 2: CIndex = 1
    Dim PrimaryKey As Variant
    For Each PrimaryKey In DataList.Keys
        
        For Each FldName In FldNames
            Sheet.Cells(RIndex, CIndex) = GetValue(PrimaryKey, FldName)
            CIndex = CIndex + 1
        Next
        RIndex = RIndex + 1
        CIndex = 1
    Next
End Sub

'デフォルトのシートからの読み込みです。
'本クラスのWriteSheetが出力したデータのほか、
'1行目がフィールド名、２行目からデータが始まるようなシートが読み込めます。
'それ以外の形式のデータを読み込みたい場合は適宜自作してください。
'プライマリキーは先頭のフィールドに自動的に設定します。
Public Sub ReadSheet(ByRef Sheet As Worksheet)
    DataList.RemoveAll
    Set FldNames = New Collection 'Initialize
    Dim RIndex, CIndex As Integer: RIndex = 1: CIndex = 1
    
    Do Until IsBlankCell(Sheet.Cells(RIndex, CIndex))
        FldNames.Add Sheet.Cells(RIndex, CIndex).Value
        CIndex = CIndex + 1
    Loop
    
    PrimaryKey = FldNames(1)
    OpenData
    RIndex = 2: CIndex = 1
    Do Until IsBlankCell(Sheet.Cells(RIndex, 1))
        If IsKey(Sheet.Cells(RIndex, 1)) Then
            Debug.Print Sheet.Cells(RIndex, 1)
        End If
        With AddNew(Sheet.Cells(RIndex, 1))
            For CIndex = 2 To FldNames.Count
                .item(FldNames(CIndex)) = Sheet.Cells(RIndex, CIndex).Value
            Next
        End With
        RIndex = RIndex + 1
    Loop
    
End Sub

'指定したフィールドを基準とした昇順/降順ソートを行ったキー配列を取得する
Public Function GetSortedKey(FieldName As Variant, Optional ASC = True) As Collection
    Dim SortKeys() As Variant
    ReDim SortKeys(0)
    Dim KeyIndex As Integer
    For KeyIndex = 0 To UBound(DataList.Keys)
        SortKeys(KeyIndex) = DataList.Keys(KeyIndex)
        ReDim Preserve SortKeys(UBound(SortKeys) + 1)
    Next
    ReDim Preserve SortKeys(UBound(SortKeys) - 1)

    QsortEx SortKeys, FieldName, 0, Size - 1
    If Not ASC Then Array_Reverse SortKeys
    Set GetSortedKey = Array2Collection_Conv(SortKeys)
End Function

'------------------------------------------------
'Private Procedures
'------------------------------------------------
Private Function RecordSetConv(Keys As Collection) As ADODB.Recordset
    Dim Rec As ADODB.Recordset
    Set Rec = New ADODB.Recordset
          
    Dim FldName As Variant
    For Each FldName In FldNames
        Rec.Fields.Append CStr(FldName), adVariant
    Next
    Rec.Open

    Dim Key As Variant
    For Each Key In Keys
        Dim Data As Scripting.Dictionary
        Set Data = DataList.item(Key)
        
        Rec.AddNew
        For Each FldName In FldNames
            Rec.Fields(FldName).Value = GetValue(Key, FldName)
        Next
        
        Rec.Update
        Rec.MoveNext
    Next
    Rec.MoveFirst
    Set RecordSetConv = Rec
End Function

'クイックソート再帰呼び出し
Private Sub QsortEx(ByRef SortKeys() As Variant, FieldName As Variant, LKeyIndex As Long, RKeyIndex As Long)
    If LKeyIndex = RKeyIndex Then Exit Sub
 
    Dim Pivot As Variant: Pivot = GetPivotValue(SortKeys, FieldName, LKeyIndex, RKeyIndex)
    
    If Pivot = "" Then Exit Sub
    Dim BorderIndex As Long
    BorderIndex = Partition(SortKeys, FieldName, LKeyIndex, RKeyIndex, Pivot)
    QsortEx SortKeys, FieldName, LKeyIndex, BorderIndex - 1    'Left
    QsortEx SortKeys, FieldName, BorderIndex, RKeyIndex         'Right

End Sub


Private Function Partition(ByRef SortKeys() As Variant, FieldName As Variant, LKeyIndex_Begin As Long, RKeyIndex_Begin As Long, Pivot As Variant) As Long
    Dim LKeyIndex As Long, RKeyIndex As Long
    LKeyIndex = LKeyIndex_Begin
    RKeyIndex = RKeyIndex_Begin
    
    Do While (LKeyIndex <= RKeyIndex)
        
        Do While (LKeyIndex <= RKeyIndex_Begin)
            If GetValue(SortKeys(LKeyIndex), FieldName) >= Pivot Then
                Exit Do
            End If
            LKeyIndex = LKeyIndex + 1
        Loop
        
        Do While (RKeyIndex > LKeyIndex_Begin)
            If Pivot > GetValue(SortKeys(RKeyIndex), FieldName) Then
                Exit Do
            End If
            RKeyIndex = RKeyIndex - 1
        Loop
        
        If LKeyIndex > RKeyIndex Then Exit Do
        Swap SortKeys, LKeyIndex, RKeyIndex
        LKeyIndex = LKeyIndex + 1
        RKeyIndex = RKeyIndex - 1
        
    Loop
    
    Partition = LKeyIndex
End Function

Private Function GetPivotValue(SortKeys() As Variant, FieldName, LKeyIndex As Long, RKeyIndex As Long) As Variant
    Dim LValue As Variant: LValue = GetValue(SortKeys(LKeyIndex), FieldName)
    
    Dim ii As Long
    For ii = LKeyIndex To RKeyIndex
        Dim RValue As Variant: RValue = GetValue(SortKeys(ii), FieldName)
        If LValue <> RValue Then
            GetPivotValue = IIf(LValue > RValue, LValue, RValue)
            Exit Function
        End If
    Next
    
    GetPivotValue = ""  'if Don't Found is null
End Function

Private Sub Swap(ByRef Ary() As Variant, LIndex As Long, RIndex As Long)
    Dim Buff As Variant
    Buff = Ary(LIndex)
    Ary(LIndex) = Ary(RIndex)
    Ary(RIndex) = Buff
End Sub

Private Sub Array_Reverse(ByRef Ary() As Variant)
    Dim Count As Integer
    Dim LIndex As Long, RIndex As Long: LIndex = 0: RIndex = UBound(Ary)
    
    While (LIndex < RIndex)
        Array_Swap Ary, LIndex, RIndex
        LIndex = LIndex + 1
        RIndex = RIndex - 1
    Wend
End Sub

'空白セル判定
Private Function IsBlankCell(ByRef TargetCell As Range) As Boolean
    IsBlankCell = True
    
    With TargetCell
        If IsEmpty(.Value) Then Exit Function
        If IsNull(.Value) Then Exit Function
        If .Value = "" Then Exit Function
    End With
    
    IsBlankCell = False
End Function
