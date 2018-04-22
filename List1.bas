Attribute VB_Name = "List1"
'глобальные переменные
    Public InputList As Worksheet
    Public OutputList As Worksheet
    Public outListName As String
' константы
    Public Const INF As String = "#EXTINF:"
    Public Const VLC As String = "#EXTVLCOPT:"
    Public Const HTTP As String = "http://"
    Public Const RTMP As String = "rtmp://"
'

Sub main()
    Call ini
    Call удал_стр
    Call валидация_данных
    Call заплнение_табл
 '    Call стр_по_столбцам аааа
 '    Call название_канала
End Sub

Sub ini()
    outListName = "m3u"
    Call создание_листа(outListName)
    Set InputList = ThisWorkbook.Sheets("Лист1")
    
    Debug.Print "ini OK"
End Sub

Function посл_запол_стр(столбец As Integer, лист As Worksheet) As Integer
    посл_запол_стр = лист.Cells(Rows.count, столбец).End(xlUp).Row
End Function

Sub удал_пустой_стр(столбец As Integer)

    ' переменные
    Dim счетчик As Integer
    
    ' инициализация
    счетчик = 0
    
    Debug.Print "удаление пустых строк...."
    
    For i = 1 To посл_запол_стр(1, InputList)
        If IsEmpty(InputList.Cells(i, столбец)) Then
            InputList.Rows(i).Delete
            счетчик = счетчик + 1
        End If
    Next i
    
    Debug.Print "удаленно " & счетчик & " пустых строк"
End Sub

Sub удал_стр()

    ' переменные
     Dim счетчик As Integer
     Dim столбец As Integer
     Dim посл_стр As Integer
     Dim диапозон As Range
     Dim строка As String
     Dim подстрока As String
     
    ' инициализация
     столбец = 1
     посл_стр = посл_запол_стр(столбец, InputList)
     Set диапозон = InputList.Range(InputList.Cells(1, столбец), InputList.Cells(посл_стр, столбец))
    
    ' цикл
        For i = посл_стр To 1 Step -1

            строка = диапозон(i, 1)
            подстрока = Mid(строка, 1, Len(VLC))

            If строка = "" Then
                диапозон(i, 1).Delete (xlShiftUp)
                счетчик = счетчик + 1
            ElseIf подстрока = VLC Then
                диапозон(i, 1).Delete (xlShiftUp)
                счетчик = счетчик + 1
            End If

        Next i

    Debug.Print "Удаленно: " & счетчик & " запесей"
End Sub

Sub валидация_данных()
    ' константы
     Const подстрока1 = "#EXT"
     Const подстрока2 = "http"
     Const подстрока3 = "rtmp"
    ' переменные
     Dim диапозон As Range
     Dim счетчик As Integer, посл_стр As Integer, столбец As Integer
     Dim строка As String, подстрока As String
     Dim f1 As Boolean, f2 As Boolean
    ' инициализация
     столбец = 1
     посл_стр = посл_запол_стр(столбец, InputList)
     Set диапозон = InputList.Range(InputList.Cells(1, столбец), InputList.Cells(посл_стр, столбец))
    
    ' цикл
    For i = посл_стр To 1 Step -1
        
        строка = диапозон(i, 1)
        подстрока = Mid(строка, 1, 4)

        If подстрока = подстрока1 And f1 = False Then
            f1 = True
            f2 = False
        ElseIf (подстрока = подстрока2 Or подстрока = подстрока3) And f2 = False Then
            f2 = True
            f1 = False
        Else
            диапозон(i, 1).Delete (xlShiftUp)
            счетчик = счетчик + 1
        End If

    Next i

    Debug.Print "при валидации удаленно: " & счетчик & " запесей"
End Sub

Sub заплнение_табл()
    ' переменные
        Dim стр_поиска      As String
        Dim стр_исходная    As String
        Dim символов        As Integer
        Dim строка          As Range
        Dim столбец         As Integer
        Dim первая_стр      As Integer 
        Dim диапозон        As Range
        Dim таблица         As Object
        Dim название_канала As String
 '        Dim таблица As Object
    ' инициализация
        стр_поиска = "#EXTINF:"
        символов = Len(стр_поиска)
        первая_стр = 1
        столбец = 1
        посл_стр = посл_запол_стр(столбец, InputList)
        Set диапозон = Range(InputList.Cells(первая_стр, столбец), InputList.Cells(посл_стр, столбец))
        
    For Each строка In диапозон
       
        стр_исходная = строка
        подстрока = Mid(стр_исходная, 1, символов)
       
        If подстрока = стр_поиска Then
            с_позиции = InStr(1, стр_исходная, ",")
            название_канала = Mid(стр_исходная, с_позиции + 1)
            Set таблица = ThisWorkbook.Worksheets("m3u").ListObjects("плэйлист").ListRows.Add
            таблица.Range(2) = название_канала
 '            строка.Interior.ColorIndex = 5
        Else
            строка.Interior.ColorIndex = 4
        End If
    Next
End Sub

Sub название_канала()
    ' переменные
        Dim стр_поиска      As String
        Dim стр_исходная    As String
        Dim позиция_сим     As Integer
        Dim подстрока       As String
        Dim столбец         As Integer
    ' инициализация
        посл_стр = посл_запол_стр(1, InputList)
        столбец = 1
        стр_поиска = ","
    
    For i = 1 To посл_стр
    
        стр_исходная = Sheets("Лист1").Cells(i, столбец)
        
        позиция_сим = InStr(1, стр_исходная, стр_поиска)
        подстрока = Mid(стр_исходная, позиция_сим + 1)
        
        Set ListRow = ThisWorkbook.Worksheets("m3u").ListObjects("плэйлист").ListRows.Add
        ListRow.Range(1) = i
        ListRow.Range(2) = подстрока
        ListRow.Range(3) = Sheets("Лист1").Cells(i, 3)
    Next i
End Sub

Sub создание_листа(имя_листа As String)
    Dim list As Worksheet
    Dim flagList As Boolean

    flagList = False

    For Each list In ActiveWorkbook.Worksheets
        If list.Name = имя_листа Then flagList = True
    Next list

    If Not flagList Then
        Set OutputList = ThisWorkbook.Sheets.Add
        OutputList.Name = имя_листа
 '        OutputList.DisplayGridlines = False
        Debug.Print "Создан лис: " & имя_листа
        Call создание_таблицы
    End If
End Sub

Sub создание_таблицы()
    ' переменные
        Dim ЛистПлэйлиста       As Worksheet
        Dim ТаблицаПлэйлиста_об As ListObject
        Dim СписокСтрок         As ListRow
        Dim счетчик             As Integer
    
    OutputList.ListObjects.Add( _
        xlSrcRange, _
        Range( _
            OutputList.Cells(1, 1), _
            OutputList.Cells(1, 5) _
            ), , _
        xlNo _
    ).Name = "плэйлист"

    'Изменяем названия граф
    OutputList.Cells(1, 1) = "id"
    OutputList.Cells(1, 1).EntireColumn.AutoFit
    OutputList.Cells(1, 2) = "Имя"
    OutputList.Cells(1, 2).EntireColumn.AutoFit
    OutputList.Cells(1, 3) = "Группа"
    OutputList.Cells(1, 3).EntireColumn.AutoFit
    OutputList.Cells(1, 4) = "Адрес"
    OutputList.Cells(1, 5) = "Дата"
    OutputList.Cells(1, 5).EntireColumn.AutoFit

End Sub

