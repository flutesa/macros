Sub tk2()
'
' Author: Burkova A.S.
' E-mail: flutesa@ya.ru
' Date:   january 2013
'
' Назначение: макрос формирования xml-кода для
' загрузки данных в техкарты и связанные с ними таблицы
' сохраняет загрузочный код в C:\xml.xml

    Dim q1, ab1 As Double
    Dim lr, lastrow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim XML As String
    Dim a, b, d, g, n, o, p, q, r, s, t, u, v, w, y, z, aa, ab, ac As String

    
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("C:\xml.xml", True)
    
    Workbooks("[макрос] inTNV.xls").Activate
    Sheets(1).Activate
    XML = ""
    
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row 'последняя строка в шаблоне
    
    For i = 2 To lastrow 'начинаем со второй, тк 1 - описание таблицы
        If Cells(i, "A").Value <> "" Then
            Sheets(1).Activate
            a = Trim(Cells(i, "A").Value) 'название техкарты
            b = Trim(Cells(i, "B").Value) 'описание техкарты
            d = Trim(Cells(i, "D").Value) 'примечание
            g = Trim(Str(Cells(i, "G").Value)) 'ид классификации
            XML = "<JOBPLAN><ORGID>РЖД</ORGID><JPNUM>" + a + "</JPNUM><DESCRIPTION>" + b + "</DESCRIPTION>" + vbCr + "<DESCRIPTION_LONGDESCRIPTION>" + d + "</DESCRIPTION_LONGDESCRIPTION>" + vbCr + "<CLASSSTRUCTUREID>0" + g + "</CLASSSTRUCTUREID>" + vbCr
            
            If Cells(i, "AA").Value <> "" Then
                t = Trim(Cells(i, "AA").Value) 'условие работы 1
                XML = XML + "<JOBPLANSPEC><SECTION/><ASSETATTRID>CONDITION</ASSETATTRID><LONGDESCRIPTION>" + t + "</LONGDESCRIPTION></JOBPLANSPEC>" + vbCr
            End If
            
            s = Trim(Str(Cells(i, "Z").Value)) 'номер техкарты как в справочнике
            XML = XML + "<JOBPLANSPEC><SECTION/><ASSETATTRID>TNKNUM</ASSETATTRID><ALNVALUE>" + s + "</ALNVALUE></JOBPLANSPEC>" + vbCr
            
            j = 0 'описание атрибутов работы
            Do While Cells(i + j, "W").Value <> ""
                If Cells(i + j, "X").Value <> "" Then
                    w = Trim(Cells(i + j, "W").Value) 'izmrab'ы (izmrab2 и пр. Tmin) наименование
                    q1 = Cells(i + j, "X").Value 'izmrab значение
                
                    If Cells(i + j, "X").Value < 1 Then 'обработка проблемы со значениями <0 (.123 получались вместо 0.123)
                        q = "0" + Trim(Str(q1))
                    ElseIf Cells(i + j, "X").Value >= 1 Then
                        q = Trim(Str(q1))
                    End If
                
                    r = UCase(Trim(Cells(i + j, "Y").Value)) 'izmrab еи
                    
                    XML = XML + "<JOBPLANSPEC><SECTION/><ASSETATTRID>" + w + "</ASSETATTRID><NUMVALUE>" + q + "</NUMVALUE><MEASUREUNITID>" + r + "</MEASUREUNITID></JOBPLANSPEC>" + vbCr
                End If
            j = j + 1
            Loop
            
            j = 0 'описание инструментов
            Do While Cells(i + j, "U").Value <> ""
                'If Cells(i + j, "V").Value = "" Then 'в случае, если у нас в соседнем столбце количество инструментов указано яно
                '    v = "1"
                'Else: v = Cells(i + j, "V").Value '" + Trim(Str(v)) + "
                'End If
                u = Trim(Cells(i + j, "U").Value)
                XML = XML + "<JOBTOOL><ITEMNUM>" + u + "</ITEMNUM><ITEMQTY>1</ITEMQTY><ITEMSETID>ТМЦ1</ITEMSETID></JOBTOOL>" + vbCr
                j = j + 1
            Loop
            
            j = 0 'описание работников
            Do While Cells(i + j, "N").Value <> "" 'специализация
                n = Trim(Cells(i + j, "N").Value)
                o = Trim(Cells(i + j, "O").Value)
                p = Trim(Cells(i + j, "P").Value)
                If Cells(i + j, "O").Value > 0 Then 'если уровень умения есть
                    XML = XML + "<JOBLABOR><CRAFT>" + n + "</CRAFT><SKILLLEVEL>" + o + "</SKILLLEVEL><QUANTITY>" + p + "</QUANTITY></JOBLABOR>" + vbCr
                Else 'If Cells(i + j, "I").Value = "" Or Cells(i + j, "O").Value = 0 Then 'если уровня умения нет или он равен 0
                    XML = XML + "<JOBLABOR><CRAFT>" + n + "</CRAFT><QUANTITY>" + p + "</QUANTITY></JOBLABOR>" + vbCr
                End If
                j = j + 1
            Loop
            
            j = 0 'операции в работе
            Do While Cells(i + j, "K").Value <> ""
                If Cells(i + j, "I").Value <> "" Then
                    y = Trim(Cells(i + j, "H").Value) 'номер операции
                    z = Trim(Cells(i + j, "I").Value) 'описание операции
                    If Cells(i + j, "H").Value = "1" Then
                        XML = XML + "<JOBTASK><JPTASK>" + y + "</JPTASK><DESCRIPTION>" + z + "</DESCRIPTION>" + vbCr
                    ElseIf Cells(i + j, "H").Value <> "1" Then
                        XML = XML + "</JOBTASK>" + vbCr + "<JOBTASK><JPTASK>" + y + "</JPTASK><DESCRIPTION>" + z + "</DESCRIPTION>" + vbCr
                    End If
                End If
                aa = Trim(Cells(i + j, "K").Value) 'атрибуты операции
                'ab1 = Cells(i + j, "L").Value 'Tuzm_nn значение
                
                If Cells(i + j, "L").Value <> "" Then
                    If Cells(i + j, "L").Value < 1 Then
                        ab = "0" + Trim(Cells(i + j, "L").Value) 'значения операции
                    ElseIf Cells(i + j, "L").Value >= 1 Then
                        ab = Trim(Cells(i + j, "L").Value) 'значения операции
                    End If
                    
                    Select Case aa 'еи операции
                        Case "IZMTASK"
                            ac = Trim(Cells(i + j, "M").Value)
                        Case "IZMTASK2"
                            ac = UCase(Trim(Cells(i + j, "M").Value))
                        Case "UCHOBRAB"
                            ac = ""
                        Case "TUZM_NN"
                            ac = "Ч-МИН"
                    End Select
                    XML = XML + "<JOBTASKSPEC><SECTION/><JPTASK>" + y + "</JPTASK><ASSETATTRID>" + aa + "</ASSETATTRID><NUMVALUE>" + ab + "</NUMVALUE><MEASUREUNITID>" + ac + "</MEASUREUNITID></JOBTASKSPEC>" + vbCr
                End If
            j = j + 1
            Loop
          
            XML = XML + "</JOBTASK></JOBPLAN>" + vbCr + vbCr + vbCr
            
            ts.Write XML
        End If
    Next

    ts.Close

End Sub
