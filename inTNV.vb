Sub tk()
'
' Author: Burkova A.S.
' E-mail: flutesa@ya.ru
' Date:   january 2013
'
' Назначение: макрос формирования шаблона для выверки
' и последующего преобразования к загрузочному коду

   
    'данные из исходных файлов
    Dim nomB1 As String         'номер ТНВ как в справочнике
    Dim idB2 As String          'ИД ТНВ
    Dim opisanB3 As String      'описание ТНВ
    Dim primB4 As String        'примечания ТНВ
    Dim inst As String          'инструменты
    Dim inststr() As String
        
    Dim izmrabeiF2 As String    'единица измерения измерителя работы
    Dim izmrabei2F3 As String   'единица измерения измерителя работы
    Dim tnorhourE4 As String    'значение измерителя работ (еи стандартно Ч-ЧАС)
    Dim yslovieE5 As String     'условие работы

    Dim Top, Tpz, Tob, Totl, TpzOo, TobOo, TotlOo, Tmin As String 'расчёт нормы времени на измеритель

    Dim opercol As Double       'кол-во операций в техкарте - для цикла
    Dim opernamB() As String    'названия операций в техкарте
    Dim izmtask() As String     'название измерителя операции
    Dim izmtask2() As String    'название измерителя операции
    Dim uchobrab() As String    'название измерителя операции
    Dim uchobrab2() As String   'название измерителя операции
    Dim tuzm_nn() As String     'значение измерителя операции (еи стандартно Ч-МИН)
    Dim izmtei() As String      'единицы измерения атрибутов работы
    Dim izmtei2() As String     'единицы измерения атрибутов работы
    
    Dim dopycol As Integer      'кол-во дополнительных условий
    Dim dopy() As String        'описание дополнительных условий
    Dim dopyvalue() As String   'значение дополнительных условий
    Dim dopyop() As String      'операция (арифметическое действие)
    
    Dim rabcol As Integer       'кол-во работников - для цикла
    Dim spezidH() As String     'ид специализации работника
    Dim skilllI() As String     'уровень умения
    Dim colrabJ() As String     'количество
     

    'даннные из Классификация.xls - функциональность пока не реализована
    'Dim idklasG1 As String      'ссылка на ID классификации
 
    Dim i, z, li As Integer      'счётчики циклов
    Dim lastrow, lr As Integer   'последняя строка по ходу формирования шаблона в выходном файле
    
    Dim noid, iddes, id As String 'для замены на ид
    
    Application.ScreenUpdating = 0
    

    '********************************************
    '********** главный цикл программы **********
    '********************************************

    'выбор необходимых файлов
    avFiles = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", , "Выберете файлы для сбора шаблона", , True)
    If VarType(avFiles) = vbBoolean Then Exit Sub
    
    aoFile = Application.GetOpenFilename("Excel files(*.xls*),*.xls*", , "Выберете выходной файлй ТехКарты.xls", , True)
    If VarType(aoFiles) = vbBoolean Then Exit Sub
    Workbooks.Open Filename:=aoFile(1)
    
    'главный цикл по книгам - кол-во выбранных техкарт из справочника
    For li = LBound(avFiles) To UBound(avFiles)
        Workbooks.Open Filename:=avFiles(li)
        
        'блок основной информации о техкарте
        nomB1 = Cells(1, "B").Value
        idB2 = Cells(2, "B").Value
        opisanB3 = Cells(3, "B").Value
        primB4 = Cells(4, "B").Value
        
        
        'инструменты
        inst = Cells(5, "B").Value
        If Right(inst, 1) = "." Then
            inst = Left(inst, Len(inst) - 1) 'удаляем точку в конце
        ElseIf Right(inst, 2) = ". " Then
            inst = Left(inst, Len(inst) - 2)
        End If
        'распарсиваем инструменты
        ReDim inststr(64) As String 'max кол-во инстр 64 шт
        inststr = Split(inst, ",")
        
        
        'блок описания атрибутов техкарты
        izmrabeiF2 = Cells(2, "F").Value
        izmrabei2F3 = Cells(3, "F").Value
        tnorhourE4 = Cells(4, "E").Value
        yslovieE5 = Cells(5, "E").Value
    
    
        'блок рассчётов нормы времени
        Top = Cells(10, "B").Value
        Tpz = Cells(10, "C").Value
        Tob = Cells(10, "D").Value
        Totl = Cells(10, "E").Value
        TpzOo = Cells(9, "C").Value
        TobOo = Cells(9, "D").Value
        TotlOo = Cells(9, "E").Value
        Tmin = Cells(10, "F").Value
        
        
        'блок операций техкарты
        opercol = Range("A13:A250").Find(Application.Max(Range("A13:A250"))).Value 'кол-во операций в техкарте
        ReDim opernamB(opercol) As String 'названия операций в техкарте
        ReDim izmtask(opercol) As String 'название измерителя операции (уточнённый объём работ на измеритель)
        ReDim izmtask2(opercol) As String 'название измерителя операции (уточнённый объём работ на измеритель)
        ReDim uchobrab(opercol) As String 'название измерителя операции (уточнённый объём работ на измеритель)
        ReDim uchobrab2(opercol) As String 'название измерителя операции (уточнённый объём работ на измеритель)
        ReDim tuzm_nn(opercol) As String 'значение измерителя операции
        ReDim izmtei(opercol) As String 'единицы измерения атрибута IZMTASK
        ReDim izmtei2(opercol) As String 'единицы измерения атрибута IZMTASK2
        'цикл формирования блока с операциями
        z = 0
        For i = 1 To opercol * 5 Step 5 ' [* 5] тк строк на каждую операцию 5 строк приходится, [Step 5] тк перескакиваем через 5 строк для следующей операции
            opernamB(z) = Cells(i + 12, "B").Value
            izmtask(z) = Cells(i + 12, "D").Value
            izmtask2(z) = Cells(i + 12 + 1, "D").Value
            uchobrab(z) = Cells(i + 12 + 2, "D").Value
            uchobrab2(z) = Cells(i + 12 + 3, "D").Value
            tuzm_nn(z) = Cells(i + 12 + 4, "D").Value
            izmtei(z) = Cells(i + 12, "E").Value
            izmtei2(z) = Cells(i + 12 + 1, "E").Value
            z = z + 1
        Next 'For i = 1 To operkol * 5 Step 5
   
   
        'блок доп. условий
        dopycol = Application.CountA(Range("I2:I10")) 'количество доп условий
        ReDim dopy(dopycol) As String
        ReDim dopyvalue(dopycol) As String
        ReDim dopyop(dopycol) As String
        z = 0
        For i = 1 To dopycol
            dopy(z) = Cells(i + 1, "H").Value ' +1 тк первая строка всегда заголовок таблицы и мы его не учитываем
            dopyvalue(z) = Cells(i + 1, "I").Value
            dopyop(z) = Cells(i + 1, "J").Value
            z = z + 1
        Next 'For i = 1 To dopycol
        
        
        'блок работников
        rabcol = Application.CountA(Range("H13:H30"))
        ReDim spezidH(rabcol) As String
        ReDim skilllI(rabcol) As String
        ReDim colrabJ(rabcol) As String
        z = 0
        For i = 1 To rabcol
            spezidH(z) = Cells(i + 12, "H").Value
            skilllI(z) = Cells(i + 12, "I").Value
            colrabJ(z) = Cells(i + 12, "J").Value
            z = z + 1
        Next 'For i = 1 To rabcol
               
        ActiveWorkbook.Close 'закрываем ткущую книгу с исходными данными
        
        '***************************************
        '* Заполнение итоговой таблицы данными *
        '***************************************
        
        Workbooks("ТехКарты.xls").Activate
        
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row 'определение "самой" последней строки по всем содержательным столбцам
        For i = 1 To 35
            lr = Cells(Rows.Count, i).End(xlUp).Row
            If lr >= lastrow Then
                lastrow = lr
            End If
        Next i 'For i = 1 To 35
        If lastrow <> 2 Then lastrow = lastrow + 1

        
        Cells(lastrow + 1, "A").Value = idB2      'ид тнв
        Cells(lastrow + 1, "B").Value = opisanB3  'описание тнв
        Cells(lastrow + 1, "D").Value = primB4    'примечаение тнв
        Range(Cells(lastrow + 1, "A"), Cells(lastrow + 1, "G")).Interior.ColorIndex = 34 'синенький, для отделения начала новой тк
        
        'Cells(lastrow + 1, "G").Formula = "=Классификация.xls!$A$3" 'ссылка на номер классификации тнв
        
        z = 0 'цикл для записи операций
        For i = 1 To opercol * 5 Step 5
            Cells(lastrow + i, "H").Value = z + 1             'номер операции в работе
            Cells(lastrow + i, "I").Value = opernamB(z)       'наименование операции
            Cells(lastrow + i, "K").Value = "IZMTASK"         'наименование атрибута измерителя операции работы
            Cells(lastrow + i, "L").Value = izmtask(z)
            Cells(lastrow + i, "M").Value = izmtei(z)
            Cells(lastrow + i + 1, "K").Value = "IZMTASK2"
            Cells(lastrow + i + 1, "L").Value = izmtask2(z)
            Cells(lastrow + i + 1, "M").Value = izmtei2(z)
            Cells(lastrow + i + 2, "K").Value = "UCHOBRAB"    'наименование атрибута измерителя операции работы
            Cells(lastrow + i + 2, "L").Value = uchobrab(z)
            Cells(lastrow + i + 3, "K").Value = "UCHOBRAB2"   'наименование атрибута измерителя операции работы
            Cells(lastrow + i + 3, "L").Value = uchobrab2(z)
            Cells(lastrow + i + 4, "K").Value = "TUZM_NN"     'наименование атрибута измерителя операции работы
            Cells(lastrow + i + 4, "L").Value = tuzm_nn(z)
            Cells(lastrow + i + 4, "M").Value = "Ч-МИН"
            z = z + 1
        Next 'For i = 1 To operkol


        z = 0 'цикл для записи работников
        For i = 1 To rabcol
            Cells(lastrow + i, "N").Value = spezidH(z)      'ид специализации работника
            Cells(lastrow + i, "O").Value = skilllI(z)      'характеристика работников
            Cells(lastrow + i, "P").Value = colrabJ(z)      'кол-во работников
            z = z + 1
        Next 'For i = 1 To rabcol
        
        
        'цикл для записи инструментов
        For i = 0 To UBound(inststr)
            Cells(lastrow + 1 + i, "U") = Trim(inststr(i))
        Next i
        
        Cells(lastrow + 1, "W").Value = "IZMRAB"        'наименование атрибута измерителя работы
        Cells(lastrow + 1, "X").Value = "1"             'значение измерителя работ
        Cells(lastrow + 1, "Y").Value = izmrabeiF2
        Cells(lastrow + 2, "W").Value = "IZMRAB2"
        If izmrabei2F3 <> "" Then Cells(lastrow + 2, "X").Value = "1"
        Cells(lastrow + 2, "Y").Value = izmrabei2F3
        Cells(lastrow + 3, "W").Value = "TNOR-HOUR"
        Cells(lastrow + 3, "X").Value = tnorhourE4
        Cells(lastrow + 3, "Y").Value = "Ч-ЧАС"
        Cells(lastrow + 4, "W").Value = "Top"
        Cells(lastrow + 4, "X").Value = Top
        Cells(lastrow + 5, "W").Value = "Tpz"
        Cells(lastrow + 5, "X").Value = Tpz
        Cells(lastrow + 6, "W").Value = "Tob"
        Cells(lastrow + 6, "X").Value = Tob
        Cells(lastrow + 7, "W").Value = "Tpotl"
        Cells(lastrow + 7, "X").Value = Totl
        Cells(lastrow + 8, "W").Value = "%Tpz"
        Cells(lastrow + 8, "X").Value = TpzOo
        Cells(lastrow + 9, "W").Value = "%Tob"
        Cells(lastrow + 9, "X").Value = TobOo
        Cells(lastrow + 10, "W").Value = "%Tpotl"
        Cells(lastrow + 10, "X").Value = TotlOo
        Cells(lastrow + 11, "W").Value = "Tmin"
        Cells(lastrow + 11, "X").Value = Tmin
        Cells(lastrow + 11, "Y").Value = "МИН"
        

        Cells(lastrow + 1, "Z").Value = nomB1        'номер техкарты в справочнике
        Cells(lastrow + 1, "AA").Value = yslovieE5   'ссылки на условия
        
        
        z = 0 'доп условия
        For i = 1 To dopycol
            Cells(lastrow + i, "AF").Value = i
            Cells(lastrow + i, "AG").Value = dopy(z)
            Cells(lastrow + i, "AH").Value = dopyvalue(z)
            Cells(lastrow + i, "AI").Value = dopyop(z)
            z = z + 1
        Next
        

        'Cells(lastrow + 1, "AE").Formula = "=Классификация.xls!$A$3" 'номер классификации справочника техкарт
        
        
        ActiveWorkbook.Save
    Next 'For li = LBound(avFiles) To UBound(avFiles) 'главный цикл по книгам - кол-во выбранных техкарт из справочника


    '**********************************
    '*** Преобразование в айдишники ***
    '**********************************
    ' требует актуальных выгрузок из БД

    'еи операций
    Workbooks("ТехКарты.xls").Activate
    lastrow = Cells(Rows.Count, "M").End(xlUp).Row
    
    Workbooks("[макрос] outTNV.xls").Activate
    Sheets("еи").Activate
    lr = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastrow 'цикл по столбцу еи в шаблоне
        Workbooks("ТехКарты.xls").Activate
        
        noid = LCase(Cells(i, "M").Value) 'текст в шаблоне
        If noid = "" Then GoTo NEXT1_
        
        Cells(i, "M").Interior.ColorIndex = 3 'красный, чё искали - не нашлося
            
            If noid = "ч-мин" Then
                Cells(i, "M").Interior.ColorIndex = 35 'зелёный
                GoTo NEXT1_
            End If
                For z = 1 To lr 'цикл по выгрузке
                    Workbooks("[макрос] outTNV.xls").Activate
                    Sheets("еи").Activate
        
                    iddes = LCase(Cells(z, "C").Value) 'описание из выгрузки
            
                    If noid = iddes Then 'если совпадают, берём ид и записываем в ячейку
                        id = Cells(z, "A").Value
                        Workbooks("ТехКарты.xls").Activate
                        Cells(i, "M").Interior.ColorIndex = 35 'зелёный
                        Cells(i, "M").Value = id
                    End If
                Next z
NEXT1_:
    Next i
    
    'еи работ
    Workbooks("ТехКарты.xls").Activate
    lastrow = Cells(Rows.Count, "Y").End(xlUp).Row
    
    Workbooks("[макрос] outTNV.xls").Activate
    Sheets("еи").Activate
    lr = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastrow 'цикл по столбцу еи в шаблоне
        Workbooks("ТехКарты.xls").Activate
        
        noid = LCase(Cells(i, "Y").Value) 'текст в шаблоне
        If noid = "" Then GoTo NEXT2_
        
        Cells(i, "Y").Interior.ColorIndex = 3 'красный, чё искали - не нашлося
            
            If noid = "ч-час" Or noid = "мин" Then
                Cells(i, "Y").Interior.ColorIndex = 35 'зелёный
                GoTo NEXT2_
            End If
                For z = 1 To lr 'цикл по выгрузке
                    Workbooks("[макрос] outTNV.xls").Activate
                    Sheets("еи").Activate
        
                    iddes = LCase(Cells(z, "C").Value) 'описание из выгрузки
            
                    If noid = iddes Then 'если совпадают, берём ид и записываем в ячейку
                        id = Cells(z, "A").Value
                        Workbooks("ТехКарты.xls").Activate
                        Cells(i, "Y").Interior.ColorIndex = 35 'зелёный
                        Cells(i, "Y").Value = id
                    End If
                Next z
NEXT2_:
    Next i
    
    'специализации
    Workbooks("ТехКарты.xls").Activate
    lastrow = Cells(Rows.Count, "N").End(xlUp).Row
    
    Workbooks("[макрос] outTNV.xls").Activate
    Sheets("специализации").Activate
    lr = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastrow 'цикл по столбцу еи в шаблоне
        Workbooks("ТехКарты.xls").Activate
        
        noid = LCase(Cells(i, "N").Value) 'текст в шаблоне
        If noid = "" Then GoTo NEXT3_
        
        Cells(i, "N").Interior.ColorIndex = 3 'красный, чё искали - не нашлося
    
            For z = 1 To lr 'цикл по выгрузке
                Workbooks("[макрос] outTNV.xls").Activate
                Sheets("специализации").Activate
        
                iddes = LCase(Cells(z, "B").Value) 'описание из выгрузки
            
                If noid = iddes Then 'если совпадают, берём ид и записываем в ячейку
                    id = Cells(z, "A").Value
                    Workbooks("ТехКарты.xls").Activate
                    Cells(i, "N").NumberFormat = "@"
                    Cells(i, "N").Interior.ColorIndex = 35 'зелёный
                    Cells(i, "N").Value = id
                End If
            Next z
NEXT3_:
    Next i
    
            
    'инструменты
    Workbooks("ТехКарты.xls").Activate
    lastrow = Cells(Rows.Count, "U").End(xlUp).Row
    
    Workbooks("[макрос] outTNV.xls").Activate
    Sheets("инструменты").Activate
    lr = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastrow 'цикл по столбцу еи в шаблоне
        Workbooks("ТехКарты.xls").Activate
        
        noid = LCase(Cells(i, "U").Value) 'текст в шаблоне
        If noid = "" Then GoTo NEXT4_
        
        Cells(i, "U").Interior.ColorIndex = 3 'красный, чё искали - не нашлося
    
            For z = 1 To lr 'цикл по выгрузке
                Workbooks("[макрос] outTNV.xls").Activate
                Sheets("инструменты").Activate
        
                iddes = LCase(Cells(z, "B").Value) 'описание из выгрузки
            
                If noid = iddes Then 'если совпадают, берём ид и записываем в ячейку
                    id = Cells(z, "A").Value
                    Workbooks("ТехКарты.xls").Activate
                    Cells(i, "U").Interior.ColorIndex = 35 'зелёный
                    Cells(i, "U").Value = id
                End If
            Next z
NEXT4_:
    Next i

Workbooks("[макрос] outTNV.xls").Activate
Sheets("макрос").Activate 'активируем заглавную вкладку после отработки макроса

Application.ScreenUpdating = 1

Workbooks("ТехКарты.xls").Activate


End Sub
