Attribute VB_Name = "通用函数"
'@venjet
'ver data 2017-8-24
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long '该声明得到系统开机到现在的时间(单位：毫秒)
Option Base 1
'返回随机加权概率数组的单次结果
Function probArrayCal(ByRef tagArray As Variant)
    Dim arrSum
    For Each element In tagArray
        arrSum = arrSum + element
    Next
    randNum = Int(1 + Rnd * arrSum)
    arrSum = 0
    cal = 0
    For Each element In tagArray
        cal = cal + 1
        arrSum = arrSum + element
        If arrSum >= randNum Then
            probCal = cal
            Exit Function
        End If
    Next
End Function
'输入随机数与精度（可选），返回是否命中
Function probNumCal(randNum, Optional accuracy = 1000) As Boolean
    If randNum >= Int(1 + Rnd * accuracy) Then
        probNumCal = True
    Else
        probNumCal = False
    End If
End Function

'输入基础概率、叠加概率、保底次数（可选，默认无）,限定次数（可选，默认无）,阈值（可选，默认10000）及精度（可选，默认万），返回实际概率
Function randProb(baseProb, addProb, Optional maxNum = 0, Optional minNum = 0, Optional threshold = 10000, Optional accuracy = 10000)
    Dim hitArr()
    Dim missArr()
    Dim probArr()
    
    '处理带有阈值的伪随机
    If threshold / accuracy < 1 Then
        baseProb = threshold
        addProb = 0
    End If
    
    '各种为0的情况
    If baseProb + addProb + maxNum = 0 Then
        randProb = 0
        Exit Function
    ElseIf addProb + maxNum + minNum = 0 Then
        randProb = baseProb / accuracy
        Exit Function
    ElseIf addProb + minNum = 0 Then
        Count = maxNum
    ElseIf addProb + maxNum = 0 Then
        For i = 1 To minNum
            fixProb = fixProb + (baseProb / accuracy) ^ i * (1 - baseProb / accuracy) ^ (minNum - i) * Application.WorksheetFunction.Combin(minNum, i)
        Next
        randProb = fixProb / minNum
        Exit Function
    ElseIf addProb = 0 Then
        Count = -Int(-Log(0.0001 * accuracy / baseProb) / Log((accuracy - baseProb) / accuracy)) + 1
    Else
        Count = -Int(-(accuracy - baseProb) / addProb)
    End If
    
    ReDim Preserve hitArr(1 To Count)
    ReDim Preserve missArr(1 To Count)
    ReDim Preserve probArr(1 To Count)

    For i = 1 To Count
        If maxNum <> 0 And i >= maxNum Then
            hitArr(i) = 1
        Else
            hitArr(i) = Application.WorksheetFunction.min((baseProb + addProb * i) / accuracy, 1)
        End If
        missArr(i) = 1 - hitArr(i)
        probArr(i) = hitArr(i)
        For j = 1 To i - 1
            probArr(i) = probArr(i) * missArr(j)
        Next j
        
        If minNum = 0 Then
            probArr(i) = i * probArr(i)
        Else
            probArr(i) = Application.WorksheetFunction.max(minNum, i) * probArr(i)
        End If
    Next
    
    randProb = 1 / Application.WorksheetFunction.Sum(probArr)
    
End Function

'使进程睡眠T毫秒
Public Function SleepToo(T As Long)
    Dim Savetime As Long
    Savetime = timeGetTime '记下开始时的时间
    While timeGetTime < Savetime + T '循环等待
        DoEvents '转让控制权
    Wend
End Function

'将传入的数组乱序
Sub randArray(ByRef tagArray As Variant)
    upbound = UBound(tagArray)
    For i = 1 To upbound
       randNum = Int(upbound * Rnd) + 1
       Temp = tagArray(i)
       tagArray(i) = tagArray(randNum)
       tagArray(randNum) = Temp
    Next i
End Sub

'返回给定区间的去重随机数组，长度可指定，默认为所有数字。
Function v_RandBetween(min As Integer, max As Integer, Optional length = 0)

    Count = max - min + 1
    
    Dim Arr()
    Dim result()
    
    ReDim Preserve Arr(Count)
    
    For i = 1 To Count
        Arr(i) = min + i - 1
    Next i
    
    Call randArray(Arr)
    
    If length = 0 Then
        length = Count
    End If
    
    ReDim Preserve result(length)
    
    For j = 1 To length
        result(j) = Arr(j)
    Next j
    
    v_RandBetween = "[" & Join(result, ",") & "]"

End Function

'将二维数组的文本转换为二维数组
Sub transStrToDoubleArray(ByVal containStr As String, ByRef Arr())
    On Error GoTo Err_Handle
    ReDim Preserve Arr(1 To 1, 1 To 4)

    containStr = Replace(containStr, "[[", "{")
    containStr = Replace(containStr, "]]", "}")
    containStr = Replace(containStr, "],[", ";")
  
    
    'Evaluate这个破函数的字符数有限制，不能超255
    'Update:听说2010以后没这限制了，开放一下试试
    'If Len(containStr) > 255 Then
    '    Debug.Print ("超过255个字符，Evaluate函数无法转换。")
    'End If
    
    Arr = Application.Evaluate(containStr)
    
    If InStr(containStr, ";") = 0 Then
         Dim ArrForDouble(1, 1 To 99999)
         For i = 1 To UBound(Arr)
            ArrForDouble(1, i) = Arr(i)
         Next
         Arr = ArrForDouble
    End If
    
    Exit Sub
Err_Handle:
    Arr = Null
End Sub

'传入二维数组文本，根据索引值返回对应数值
Function v_GetDArrayValue(containStr As String, COl As Integer, row As Integer)
    On Error GoTo Err_Handle
    Dim Arr()
    Call transStrToDoubleArray(containStr, Arr)
    v_GetDArrayValue = Arr(COl, row)
    Exit Function
Err_Handle:
    v_GetDArrayValue = ""
End Function

'传入二维数组文本，根据索引值返回对应列的一维数组文本
Function v_GetDArrayCol(containStr As String, COl As Integer)
    On Error GoTo Err_Handle
    Dim Arr()
    Dim result()
    Call transStrToDoubleArray(containStr, Arr)
    ReDim Preserve result(UBound(Arr, 1) - 1)
    For i = 1 To UBound(Arr, 1)
            result(i - 1) = Arr(i, COl)
    Next
    v_GetDArrayCol = "[" & Join(result, ",") & "]"
    Exit Function
Err_Handle:
    v_GetDArrayCol = ""
End Function

'传入二维数组文本，根据索引值返回对应行的一维数组文本
Function v_GetDArrayRow(containStr As String, row As Integer)
    On Error GoTo Err_Handle
    Dim Arr()
    Dim result()
    Call transStrToDoubleArray(containStr, Arr)
    ReDim Preserve result(UBound(Arr, 2) - 1)
    For i = 1 To UBound(Arr, 2)
            result(i - 1) = Arr(row, i)
    Next
    v_GetDArrayRow = "[" & Join(result, ",") & "]"
    Exit Function
Err_Handle:
    v_GetDArrayRow = ""
End Function

'将一维数组的文本转换为一维数组
Sub transStrToSingleArray(ByVal containStr As String, ByRef Arr)
    On Error GoTo Err_Handle
    'ReDim Preserve arr(1 To 1, 1 To 4)

    containStr = Replace(containStr, "[", "")
    containStr = Replace(containStr, "]", "")
    
    Arr = Split(containStr, ",")
    
    Exit Sub
Err_Handle:
    Arr = Null
End Sub

'传入一维数组文本，根据索引值返回对应数值
Function v_GetSArrayValue(containStr As String, COl As Integer)
    On Error GoTo Err_Handle
    Dim Arr As Variant
    Call transStrToSingleArray(containStr, Arr)
    v_GetSArrayValue = Arr(COl - 1)
    Exit Function
Err_Handle:
    v_GetSArrayValue = ""
End Function

'将某列数据进行乘法及加法取整后，转换为一维数组文本形式(默认乘1加0)
Function v_ColToArrStr(rangeArr As Range, Optional a = 1, Optional b = 0, Optional isRound = 0)
    Dim Arr
    Arr = Application.Transpose(rangeArr)
    If a = 1 And b = 0 And isRound = 0 Then
        v_ColToArrStr = "[" & Join(Arr, ",") & "]"
        Exit Function
    End If
    For i = 1 To UBound(Arr)
        If isRound = 1 Then
            Arr(i) = Round(Arr(i) * a + b, 0)
        Else
            Arr(i) = Arr(i) * a + b
        End If
    Next
    v_ColToArrStr = "[" & Join(Arr, ",") & "]"
End Function

'将某行数据进行乘法及加法后，转换为一维数组文本形式
Function v_RowToArrStr(rangeArr As Range, Optional a = 1, Optional b = 0, Optional isRound = 0)
    Dim Arr
    Arr = Application.Transpose(Application.Transpose(rangeArr))
    If a = 1 And b = 0 And isRound = 0 Then
        v_RowToArrStr = "[" & Join(Arr, ",") & "]"
        Exit Function
    End If
    For i = 1 To UBound(Arr)
        If isRound = 1 Then
            Arr(i) = Round(Arr(i) * a + b, 0)
        Else
            Arr(i) = Arr(i) * a + b
        End If
    Next
    v_RowToArrStr = "[" & Join(Arr, ",") & "]"
End Function

'输入价值、数量、概率数组，返回总价值
Function v_valueArrSum(valueArrStr, numArrStr, probArrStr)
    On Error GoTo Err_Handle
    Dim valueArr As Variant
    Dim numArr As Variant
    Dim probArr As Variant
    
    Call transStrToSingleArray(valueArrStr, valueArr)
    Call transStrToSingleArray(numArrStr, numArr)
    Call transStrToSingleArray(probArrStr, probArr)
    For i = 0 To UBound(valueArr)
        v_valueArrSum = v_valueArrSum + (valueArr(i) + 0) * Int(numArr(i)) * Int(probArr(i)) / 1000
    Next
    Exit Function
Err_Handle:
    v_valueArrSum = 0
End Function


'输入一维数组文本，返回总值
Function v_sArrSum(valueArrStr)
    On Error GoTo Err_Handle
    Dim valueArr As Variant
 
    Call transStrToSingleArray(valueArrStr, valueArr)

    For i = 0 To UBound(valueArr)
        v_sArrSum = v_sArrSum + (valueArr(i) + 0)
    Next
    Exit Function
Err_Handle:
    v_sArrSum = 0
End Function

'输入基础值，长度，等差值（默认0），等比值（默认1），返回等差等比数组文本
Function v_AriGeoArray(base, length, Optional ari = 0, Optional geo = 1)
    For i = 0 To length - 1
        v_AriGeoArray = v_AriGeoArray & "," & (base + base * (geo - 1) * i + ari * i)
    Next
    
    strLen = Len(v_AriGeoArray)
    
    v_AriGeoArray = "[" & Mid(v_AriGeoArray, 2, strLen) & "]"

End Function


'返回某列数据的平方平均数
Function v_ColRMSquare(rangeArr As Range)
    Dim Arr
    Arr = Application.Transpose(rangeArr)
    For i = 1 To UBound(Arr)
        arrSum = arrSum + Arr(i) ^ 2
    Next
    v_ColRMSquare = (arrSum / UBound(Arr)) ^ 0.5
End Function
'返回某行数据的平方平均数
Function v_RowRMSquare(rangeArr As Range)
    Dim Arr
    Arr = Application.Transpose(Application.Transpose(rangeArr))
    For i = 1 To UBound(Arr)
        arrSum = arrSum + Arr(i) ^ 2
    Next
    v_RowRMSquare = (arrSum / UBound(Arr)) ^ 0.5
End Function

'输入A,B两个数组文本，确定A是否为B数组的子集
Function v_isSubArr(childStr As String, fatherStr As String)
    
    If Len(childStr) = 0 Then
        v_isSubArr = 0
        Exit Function
    End If
    
    Dim childArr
    Dim fatherArr
    Call transStrToSingleArray(childStr, childArr)
    Call transStrToSingleArray(fatherStr, fatherArr)
    
    Count = 0
    
    For i = 0 To UBound(childArr)
        For j = 0 To UBound(fatherArr)
            If childArr(i) = fatherArr(j) Then
                Count = Count + 1
            End If
        Next j
    Next i
    
    If Count < UBound(childArr) + 1 Or Len(childStr) = 0 Then
        v_isSubArr = 0
    Else
        v_isSubArr = 1
    End If
End Function

'@venjet
'输入数字，保留前几位（可选，默认2位），多少以下直接取整（可选，默认100）
Function v_cutNum(beCut, Optional retainNum = 99, Optional precision = 100)

    beCut = Int(beCut)

    '尾部处理
    If beCut < precision Then
        precision = 10 ^ Int(Len(beCut) - 1)
        v_cutNum = beCut
    Else
        v_cutNum = Application.Round(beCut / precision, 0) * precision
    End If
    
    '头部处理
    headNum = Int(Mid(v_cutNum, 1, retainNum)) '从头开始截取的数字
    digitsNum = Application.max(0, Int(Len(v_cutNum) - retainNum)) '剩余的位数
   
    v_cutNum = headNum * 10 ^ digitsNum
    
End Function

'将unix 时间戳转换为时间，注意要将单元格格式改为日期或时间
Function v_stampToTime(stamp)
    v_stampToTime = (stamp + 8 * 3600) / 86400 + 70 * 365 + 19
End Function


'简单写写的Json格式解析函数，凑活够用
'暂不支持数组value，暂不支持Json嵌套
'@venjet
'输入  jsonStr:待解析的Json语句；key:需查找的key值
'输出  对应的value
Function v_getJsonValue(jsonStr As String, key As String)
    On Error GoTo Err_Handle
        starNum = Application.Find(key, jsonStr) + Len(key) + 1
        endNum = Application.Find(",", jsonStr, starNum)
        v_getJsonValue = Mid(jsonStr, starNum, endNum - starNum)
    Exit Function
Err_Handle:
    v_getJsonValue = ""
End Function


'获取某个数组的维度
Function v_getArrarDimensions(Arr)

      On Error GoTo FinalDimension
      
      For DimNum = 1 To 60000
         'It is necessary to do something with the LBound to force it
         'to generate an error.
         ErrorCheck = LBound(Arr, DimNum)
      Next DimNum

      Exit Function

      ' The error routine.
FinalDimension:
        v_getArrarDimensions = DimNum - 1
        
End Function

'vba傻逼补丁第一弹
'vba的白痴Transpose函数不知道为什么不支持单格超256个字符
'没办法只好自己写一个了，简直了
'目前最多支持到二维...话说本来好像也就到二维吧
'另外数组从1开始，设置了从零开始的话会丢数据哦
Sub v_transpose(ByRef Arr)
    Dim newArr
    arrD = v_getArrarDimensions(Arr)
    On Error GoTo Err_Handle '这里是用来处理单一长度数组的临时方案，等外部处理完了要删掉。
    arrX = UBound(Arr)
    arrY = UBound(Arr, arrD)
    If arrD = 1 Then
        ReDim newArr(1 To arrY, 1)
        For i = 1 To arrY
            newArr(i, 1) = Arr(i)
        Next i
    ElseIf arrY = 1 Then
        ReDim newArr(1 To arrX)
        For i = 1 To arrX
            newArr(i) = Arr(i, 1)
        Next i
    Else
        ReDim newArr(1 To arrY, 1 To arrX)
        For i = 1 To arrX
            For j = 1 To arrY
                newArr(j, i) = Arr(i, j)
            Next j
        Next i
    
    End If
    
    Arr = newArr
    Exit Sub
Err_Handle:
    'Debug.Print ("v_transpose Error")  其实也不能算错...先注了吧
    ReDim newArr(1)
    newArr(1) = Arr
    Arr = newArr
End Sub

'删除一个数组中的某一行或某一列数据
Function v_DelLineOrCol(Arr As Variant, Optional delLine As Long, Optional delCol As Long)
    Dim Line As Long
    Dim lstLine As Long
    Dim COl As Long
    Dim lstCol As Long
    Dim arrNew() As Variant
    Dim Tmp  As Boolean
    lstLine = UBound(Arr, 1)
    lstCol = UBound(Arr, 2)
    '对行进行处理
    If delLine > 0 Then
        ReDim arrNew(1 To lstLine - 1, 1 To lstCol)
        For Line = 1 To lstLine
            If Line = delLine Then
                Tmp = True
            Else
                For COl = 1 To lstCol
                    arrNew(Line + Tmp, COl) = Arr(Line, COl)
                Next COl
            End If
        Next Line
        Del_LineOrCol = arrNew
        Exit Function
    End If
    '对列进行处理
    If delCol > 0 Then
        ReDim arrNew(1 To lstLine, 1 To lstCol - 1)
        For COl = 1 To lstCol
            If COl = delCol Then
                Tmp = True
            Else
                For Line = 1 To lstLine
                    arrNew(Line, COl + Tmp) = Arr(Line, COl)
                Next Line
            End If
        Next COl
        Del_LineOrCol = arrNew
    End If
End Function

'vba傻逼补丁第二弹
'用以解决filter居然不带精确匹配的问题
'返回值和原来的filter一样，但是只能返回包含的数组，多了个选项决定是否精确匹配
'值得注意的是，无论 Option Base怎么设置，数组索引值都是从0开始（要怪就怪split函数...）
Function v_Filter(myArray, myMatch, Optional isExactly = 1)
    'myMarker和myDelimiter必须是字符
    '且该字符不会出现在数组的任何元素中!
    Const myMarker As String = "♂"
    Const myDelimiter As String = "♀"
    Dim mySearchArray As Variant
    Dim myFilteredArray As Variant
    
    myFilteredArray = Filter(myArray, myMatch)
 
    If UBound(myFilteredArray) > -1 And isExactly = 1 Then
    
        '标记每个找到的元素的开始和结束
        mySearchArray = Split(myMarker & Join(myFilteredArray, myMarker & myDelimiter & myMarker) & myMarker, myDelimiter)
        '下面筛选修改后的数组
        myFilteredArray = Filter(mySearchArray, myMarker & myMatch & myMarker)
        '从结果中移除标记
        myFilteredArray = Split(Replace(Join(myFilteredArray, myDelimiter), myMarker, ""), myDelimiter)
    End If
    v_Filter = myFilteredArray
    
End Function

'将二维数组的某一行赋值给一维数组
'通过结合v_transpose函数应该能将列也进行赋值，虽然应该不太常用
Sub v_getDoubleArrayRow(ByRef sArr, dArr, rowNum)
'用错误处理来解决index函数的限制，强行写循环
On Error GoTo loopJump
    sArr = WorksheetFunction.index(dArr, rowNum, 0)
    Exit Sub
loopJump:
    ReDim sArr(UBound(dArr, 2))
    For i = 1 To UBound(dArr, 2)
        sArr(i) = dArr(rowNum, i)
    Next i
End Sub



'获取某个内容在某列（默认B列）的行号
'一句话函数系列= =
'@venjet
Function getContentRow(content, Optional COl = "B:B")
    On Error GoTo Err_Handle
        getContentRow = WorksheetFunction.Match(content, Range(COl), 0)
        Exit Function
Err_Handle:
        getContentRow = 0
        'MsgBox ("getContentRow未能找到对应内容")
End Function


'作用如函数名，判断某个路径是否存在
Public Function FileFolderExists(strFullPath As String) As Boolean

    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0

End Function

'根据给定的路径和文件名导入CSV
'注意：必须是逗号分隔
'@venjet
Sub inputCSV(path As String, name As String)
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & path & name & ".csv", Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveWorkbook.Connections(name).Delete
End Sub


