Attribute VB_Name = "ͨ�ú���"
'@venjet
'ver data 2019-12-16
Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long '�������õ�ϵͳ���������ڵ�ʱ��(��λ������)
Option Base 1
'���������Ȩ��������ĵ��ν��
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
'����������뾫�ȣ���ѡ���������Ƿ�����
Function probNumCal(randNum, Optional accuracy = 1000) As Boolean
    If randNum >= Int(1 + Rnd * accuracy) Then
        probNumCal = True
    Else
        probNumCal = False
    End If
End Function

'����������ʡ����Ӹ��ʡ����״�������ѡ��Ĭ���ޣ�,�޶���������ѡ��Ĭ���ޣ�,��ֵ����ѡ��Ĭ��10000�������ȣ���ѡ��Ĭ���򣩣�����ʵ�ʸ���
Function randProb(baseProb, addProb, Optional maxNum = 0, Optional minNum = 0, Optional threshold = 10000, Optional accuracy = 10000)
    Dim hitArr()
    Dim missArr()
    Dim probArr()
    
    '���������ֵ��α���
    If threshold / accuracy < 1 Then
        baseProb = threshold
        addProb = 0
    End If
    
    '����Ϊ0�����
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


'����һ��Ŀ�������Ŀ�������и���ΪPn����������ΪNn������£���Ҫ���ٴβ���ȫ�����С�
'һ������ڼ�����׵���Ʒ������Ҫ�ռ��Ĵ��������߿������ڼ���Ť�����ҵ�ĳ�����
'���룺���ΧRange����ʽҪ��Ϊ�������ݣ�ǰһ��Ϊ���ʣ���һ��Ϊ����������
'���磺A��    B��
'      0.3     1
'      0.2     2
'      0.4     1
'�����ȫ�����е�����������

Function v_expectTimes(probRange As Range, numRange As Range)
    Dim probArr()
    Dim numArr()
    
    probArr = Application.Transpose(probRange)
    numArr = Application.Transpose(numRange)
    
    v_expectTimes = expectRescursion(probArr, numArr)
    
End Function

'���������������Ǹ������ĺ��ĵݹ�
Function expectRescursion(probArr As Variant, numArr As Variant)
    For i = 1 To UBound(numArr)
        If numArr(i) <= 0 Then
            probArr = v_DelElemInArray(probArr, i)
            numArr = v_DelElemInArray(numArr, i)
            Exit For
        End If
    Next i
    
    arrLength = UBound(numArr)
    
    If arrLength = 1 Then
        expectRescursion = numArr(1) / probArr(1)
    Else
        expectRescursion = 1 / Application.Sum(probArr)
        For Index = 1 To arrLength
            Dim rProbArr
            Dim rNumArr
            rProbArr = probArr
            rNumArr = numArr
            rNumArr(Index) = rNumArr(Index) - 1
            expectRescursion = expectRescursion + probArr(Index) * expectRescursion(rProbArr, rNumArr) / Application.Sum(probArr)
        Next Index
    End If
End Function

'ɾ��һά�����ĳ��ֵ
Function v_DelElemInArray(tagArray As Variant, posi)
    newLength = UBound(tagArray) - 1
    
    For i = posi To newLength
        tagArray(i) = tagArray(i + 1)
    Next i
    
    ReDim Preserve tagArray(newLength)
    
    v_DelElemInArray = tagArray
    
End Function
    

'ɾ��һ�������е�ĳһ�л�ĳһ������
Function v_DelLineOrCol(Arr As Variant, Optional delLine As Long, Optional delCol As Long)
    Dim Line As Long
    Dim lstLine As Long
    Dim COl As Long
    Dim lstCol As Long
    Dim arrNew() As Variant
    Dim Tmp  As Boolean
    lstLine = UBound(Arr, 1)
    lstCol = UBound(Arr, 2)
    '���н��д���
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
    '���н��д���
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


'ʹ����˯��T����
Public Function SleepToo(T As Long)
    Dim Savetime As Long
    Savetime = timeGetTime '���¿�ʼʱ��ʱ��
    While timeGetTime < Savetime + T 'ѭ���ȴ�
        DoEvents 'ת�ÿ���Ȩ
    Wend
End Function

'���������������
Sub randArray(ByRef tagArray As Variant)
    upbound = UBound(tagArray)
    For i = 1 To upbound
       randNum = Int(upbound * Rnd) + 1
       temp = tagArray(i)
       tagArray(i) = tagArray(randNum)
       tagArray(randNum) = temp
    Next i
End Sub

'���ظ��������ȥ��������飬���ȿ�ָ����Ĭ��Ϊ�������֡�
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

'����ά������ı�ת��Ϊ��ά����
Sub transStrToDoubleArray(ByVal containStr As String, ByRef Arr())
    On Error GoTo Err_Handle
    ReDim Preserve Arr(1 To 1, 1 To 4)

    containStr = Replace(containStr, "[[", "{")
    containStr = Replace(containStr, "]]", "}")
    containStr = Replace(containStr, "],[", ";")
  
    
    'Evaluate����ƺ������ַ��������ƣ����ܳ�255
    'Update:��˵2010�Ժ�û�������ˣ�����һ������
    'If Len(containStr) > 255 Then
    '    Debug.Print ("����255���ַ���Evaluate�����޷�ת����")
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

'�����ά�����ı�����������ֵ���ض�Ӧ��ֵ
Function v_GetDArrayValue(containStr As String, COl As Integer, row As Integer)
    On Error GoTo Err_Handle
    Dim Arr()
    Call transStrToDoubleArray(containStr, Arr)
    v_GetDArrayValue = Arr(COl, row)
    Exit Function
Err_Handle:
    v_GetDArrayValue = ""
End Function

'�����ά�����ı�����������ֵ���ض�Ӧ�е�һά�����ı�
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

'�����ά�����ı�����������ֵ���ض�Ӧ�е�һά�����ı�
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

'��һά������ı�ת��Ϊһά����
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

'����һά�����ı�����������ֵ���ض�Ӧ��ֵ
Function v_GetSArrayValue(containStr As String, COl As Integer)
    On Error GoTo Err_Handle
    Dim Arr As Variant
    Call transStrToSingleArray(containStr, Arr)
    v_GetSArrayValue = Arr(COl - 1)
    Exit Function
Err_Handle:
    v_GetSArrayValue = ""
End Function

'��ĳ�����ݽ��г˷����ӷ�ȡ����ת��Ϊһά�����ı���ʽ(Ĭ�ϳ�1��0)
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

'��ĳ�����ݽ��г˷����ӷ���ת��Ϊһά�����ı���ʽ
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

'�����ֵ���������������飬�����ܼ�ֵ
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


'����һά�����ı���������ֵ
Function v_sArrSum(valueArrStr)
    On Error GoTo Err_Handle
    Dim valueArr As Variant
 
    Call transStrToSingleArray(valueArrStr, valueArr)
    
    'ת���������������ı����ͣ�+0����ǿ������ת��
     For i = 0 To UBound(valueArr)
        v_sArrSum = v_sArrSum + (valueArr(i) + 0)
    Next
    
    Exit Function
Err_Handle:
    v_sArrSum = 0
    
End Function

'�������ֵ�����ȣ��Ȳ�ֵ��Ĭ��0�����ȱ�ֵ��Ĭ��1�������صȲ�ȱ������ı�
Function v_AriGeoArray(base, length, Optional ari = 0, Optional geo = 1)
    For i = 0 To length - 1
        v_AriGeoArray = v_AriGeoArray & "," & (base + base * (geo - 1) * i + ari * i)
    Next
    
    strLen = Len(v_AriGeoArray)
    
    v_AriGeoArray = "[" & Mid(v_AriGeoArray, 2, strLen) & "]"

End Function


'����ĳ�����ݵ�ƽ��ƽ����
Function v_ColRMSquare(rangeArr As Range)
    Dim Arr
    Arr = Application.Transpose(rangeArr)
    For i = 1 To UBound(Arr)
        arrSum = arrSum + Arr(i) ^ 2
    Next
    v_ColRMSquare = (arrSum / UBound(Arr)) ^ 0.5
End Function
'����ĳ�����ݵ�ƽ��ƽ����
Function v_RowRMSquare(rangeArr As Range)
    Dim Arr
    Arr = Application.Transpose(Application.Transpose(rangeArr))
    For i = 1 To UBound(Arr)
        arrSum = arrSum + Arr(i) ^ 2
    Next
    v_RowRMSquare = (arrSum / UBound(Arr)) ^ 0.5
End Function

'����A,B���������ı���ȷ��A�Ƿ�ΪB������Ӽ�
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
'�������֣�����ǰ��λ����ѡ��Ĭ��2λ������������ֱ��ȡ������ѡ��Ĭ��100��
Function v_cutNum(beCut, Optional retainNum = 99, Optional precision = 100)

    beCut = Int(beCut)

    'β������
    If beCut < precision Then
        precision = 10 ^ Int(Len(beCut) - 1)
        v_cutNum = beCut
    Else
        v_cutNum = Application.Round(beCut / precision, 0) * precision
    End If
    
    'ͷ������
    headNum = Int(Mid(v_cutNum, 1, retainNum)) '��ͷ��ʼ��ȡ������
    digitsNum = Application.max(0, Int(Len(v_cutNum) - retainNum)) 'ʣ���λ��
   
    v_cutNum = headNum * 10 ^ digitsNum
    
End Function

'��unix ʱ���ת��Ϊʱ�䣬ע��Ҫ����Ԫ���ʽ��Ϊ���ڻ�ʱ��
Function v_stampToTime(stamp)
    v_stampToTime = (stamp + 8 * 3600) / 86400 + 70 * 365 + 19
End Function


'��дд��Json��ʽ�����������ջ��
'�ݲ�֧������value���ݲ�֧��JsonǶ��
'@venjet
'����  jsonStr:��������Json��䣻key:����ҵ�keyֵ
'���  ��Ӧ��value
Function v_getJsonValue(jsonStr As String, key As String)
    On Error GoTo Err_Handle
        starNum = Application.Find(key, jsonStr) + Len(key) + 1
        endNum = Application.Find(",", jsonStr, starNum)
        v_getJsonValue = Mid(jsonStr, starNum, endNum - starNum)
    Exit Function
Err_Handle:
    v_getJsonValue = ""
End Function


'��ȡĳ�������ά��
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

'vbaɵ�Ʋ�����һ��
'vba�İ׳�Transpose������֪��Ϊʲô��֧�ֵ���256���ַ�
'û�취ֻ���Լ�дһ���ˣ���ֱ��
'Ŀǰ���֧�ֵ���ά...��˵��������Ҳ�͵���ά��
'���������1��ʼ�������˴��㿪ʼ�Ļ��ᶪ����Ŷ
Sub v_transpose(ByRef Arr)
    Dim newArr
    arrD = v_getArrarDimensions(Arr)
    On Error GoTo Err_Handle '��������������һ�����������ʱ���������ⲿ��������Ҫɾ����
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
    'Debug.Print ("v_transpose Error")  ��ʵҲ�������...��ע�˰�
    ReDim newArr(1)
    newArr(1) = Arr
    Arr = newArr
End Sub



'vbaɵ�Ʋ����ڶ���
'���Խ��filter��Ȼ������ȷƥ�������
'����ֵ��ԭ����filterһ��������ֻ�ܷ��ذ��������飬���˸�ѡ������Ƿ�ȷƥ��
'ֵ��ע����ǣ����� Option Base��ô���ã���������ֵ���Ǵ�0��ʼ��Ҫ�־͹�split����...��
Function v_Filter(myArray, myMatch, Optional isExactly = 1)
    'myMarker��myDelimiter�������ַ�
    '�Ҹ��ַ����������������κ�Ԫ����!
    Const myMarker As String = "��"
    Const myDelimiter As String = "��"
    Dim mySearchArray As Variant
    Dim myFilteredArray As Variant
    
    myFilteredArray = Filter(myArray, myMatch)
 
    If UBound(myFilteredArray) > -1 And isExactly = 1 Then
    
        '���ÿ���ҵ���Ԫ�صĿ�ʼ�ͽ���
        mySearchArray = Split(myMarker & Join(myFilteredArray, myMarker & myDelimiter & myMarker) & myMarker, myDelimiter)
        '����ɸѡ�޸ĺ������
        myFilteredArray = Filter(mySearchArray, myMarker & myMatch & myMarker)
        '�ӽ�����Ƴ����
        myFilteredArray = Split(Replace(Join(myFilteredArray, myDelimiter), myMarker, ""), myDelimiter)
    End If
    v_Filter = myFilteredArray
    
End Function

'����ά�����ĳһ�и�ֵ��һά����
'ͨ�����v_transpose����Ӧ���ܽ���Ҳ���и�ֵ����ȻӦ�ò�̫����
Sub v_getDoubleArrayRow(ByRef sArr, dArr, rowNum)
'�ô����������index���������ƣ�ǿ��дѭ��
On Error GoTo loopJump
    sArr = WorksheetFunction.Index(dArr, rowNum, 0)
    Exit Sub
loopJump:
    ReDim sArr(UBound(dArr, 2))
    For i = 1 To UBound(dArr, 2)
        sArr(i) = dArr(rowNum, i)
    Next i
End Sub

'��ȡĳ��������ĳ�У�Ĭ��B�У����к�
'һ�仰����ϵ��= =
'@venjet
Function getContentRow(content, Optional COl = "B:B")
    On Error GoTo Err_Handle
        getContentRow = WorksheetFunction.Match(content, Range(COl), 0)
        Exit Function
Err_Handle:
        getContentRow = 0
        'MsgBox ("getContentRowδ���ҵ���Ӧ����")
End Function

'�����纯�������ж�ĳ��·���Ƿ����
Public Function FileFolderExists(strFullPath As String) As Boolean

    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0

End Function

'�ж��ļ��Ƿ����  ����"D:\ѧϰ����\2019\��������.avi"
Private Function FileExists(fname) As Boolean
    Dim x As String
    x = Dir(fname)
    If x <> "" Then FileExists = True _
        Else FileExists = False
End Function

'��ȡ·���е��ļ���
Public Function FileNameOnly(pname) As String
    Dim temp As Variant
    temp = Split(pname, Application.PathSeparator)
    FileNameOnly = temp(UBound(temp))
End Function

'�жϹ��������Ƿ����������
Private Function SheetExists(sname) As Boolean
    Dim x As Object
    On Error Resume Next
    Set x = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else: SheetExists = False
End Function

'�жϹ������Ƿ��
Private Function WorkbookIsOpen(wbname) As Boolean
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function

'��ȡ�رյĹ�������ĳ����Ԫ������ݣ���Ϊ�Ǻ꺯��(?)��ֻ����vba�����е��ã������ڱ����ֱ�ӷ���ֵ��
'path-�ļ�·����D:\Config\  or  ThisWorkbook.path
'file-�ļ�����item.xlsm...
'sheet-������sheet2 or item...
'ref-��Ԫ��������C3 or Cells(r,c).Address��������ʽ������forѭ����ȡĳ����������ݣ�
Function getClosedSheetValue(path, file, sheet, ref)
    Dim arg As String
    'TODO�����������еĺ�������ȷ�ж��ļ��Ƿ����
    If Right(path, 1) <> "\" Then path = path & "\"
    If Dir(path & file) = "" Then
        GetValue = "File not Found"
        Exit Function
    End If
    arg = "'" & path & "[" & file & "]" & sheet & "'!" & Range(ref).Range("A1").Address(, , xlR1C1)
    getClosedSheetValue = ExecuteExcel4Macro(arg)
End Function


'���ݸ�����·�����ļ�������CSV
'ע�⣺�����Ƕ��ŷָ�
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

'�ṩ�ļ�·��������Ӧ�úͱ����󣬴򿪱��
'��������1.Dim myApp As New Application
'��������2.Dim excelBook As Workbook
Sub openExcel(fullPath, ByRef myApp As Application, ByRef excelBook As Workbook)
    
    myApp.Visible = False
    
    'If excelBook.name = "" Then
    On Error GoTo File_Err_Handle '�޴��ļ�
        Set excelBook = myApp.Workbooks.Open(fullPath)
    'End If
    
    Exit Sub
File_Err_Handle:
    myApp.DisplayAlerts = False
    excelBook.Close savechanges:=False
    Set excelBook = Nothing
    myApp.Quit
    MsgBox ("File not found:  " & fileName)
    
End Sub

'����Ӧ�úͱ����󣬹رձ��
Sub closeExcel(ByRef myApp As Application, ByRef excelBook As Workbook)
    On Error GoTo File_Err_Handle

    If myApp.name <> "" Then
        myApp.DisplayAlerts = False
        excelBook.Close savechanges:=False
        Set excelBook = Nothing
        myApp.Quit
    Else
        MsgBox ("Sheets not found.")
    End If
    
    Exit Sub
    
File_Err_Handle:
    MsgBox ("Sheets not found.")
    
End Sub

'####################��������####################
'====================��������====================
'��1�������ַ���
'\\����ʾת���ַ���\����
'\t����ʾһ����\t�����ţ�
'\n��ƥ�任�У�\n�����ţ�
'
'��2���ַ�����
'[abc]����ʾ�������ַ�a���ַ�b���ַ�c�е�����һλ��
'[^abc]����ʾ�����ַ�a��b��c�е�����һλ��
'[a-z]�����е�Сд��ĸ��
'[a-zA-Z]����ʾ���е���ĸ��
'[0-9]����ʾ�����һλ���֣�
'[һ-��]����ʾ�����һ�����֣�
'
'��3���򻯵��ַ������ʽ��
'.��һ���㣬��ʾ�����һλ�ַ���
'\d���ȼ��ڡ�[0-9]�������ڼ�д����
'\D���ȼ��ڡ�[^0-9]��,���ڼ�д����
'\s����ʾ����Ŀհ��ַ������磺��\t����\n��
'\S����ʾ����ķǿհ��ַ���
'\w���ȼ��ڡ�[a-zA-Z_0-9]������ʾ���������ĸ�����֡��»�����ɣ�
'\W���ȼ��ڡ�[^a-zA-Z_0-9]������ʾ�������������ĸ�����֡��»�����ɣ�
'
'��4���߽�ƥ��
'^������Ŀ�ʼ��
'$������Ľ�����
'
'��5���������
'����?����ʾ��������Գ���0�λ�1�Σ�����\d?����ʾ����0�λ�1�����֣�
'����+����ʾ��������Գ���1�λ�1�����ϣ�
'����*����ʾ��������Գ���0�Ρ�1�λ��Σ�
'����{n}����ʾ���������ó���n�Σ�
'����{n,}����ʾ���������n�����ϣ�
'����{n,m}����ʾ���������n~m�Σ�
'
'��6���߼�����
'����1����2������1�ж���ɺ�����ж�����2��
'����1|����2������1��������2��һ�����㼴�ɣ�
'
'
'����������ʽ
'
'1��ƥ���ʱ࣬�ʱ���6λ���֡�������ʽ��\d{6}
'2��ƥ���ֻ����ֻ�����11λ���֡�������ʽ��\d{11}
'3��ƥ��绰���绰������-������ɣ�������3��4λ��������6��9λ��������ʽ��\d{3,4}-\d{6,9}
'4��ƥ�����ڣ����ڸ�ʽ��1992-5-30���������ּӺ�����ɡ�������ʽ��\d{4}-\d{1,2}-\d{1,2}



'�ж��ַ����Ƿ����������ʽ�����ط��ϱ��ʽ������
'oriText��Ŀ���ַ�����patternText��������ʽ
'isGlobal��True-ƥ�����У�False-ƥ���һ�������Ĭ��ƥ������
'isIgnoreCase��True-�����ִ�Сд��False-���ִ�Сд��Ĭ�ϲ����ִ�Сд
Function v_regexCheck(oriText, patternText, Optional isGlobal = True, Optional isIgnoreCase = True)
    With CreateObject("VBscript.regexp")
        .Pattern = patternText
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        v_regexCheck = .Execute(oriText).Count
    End With
End Function

'���������ȡ����������ʽ�����ݣ�Ĭ����ȡ��һ����û�з��ϵ����ݣ�����""
'oriText��Ŀ���ַ�����patternText��������ʽ
'isGlobal��True-ƥ�����У�False-ƥ���һ�������Ĭ��ƥ������
'isIgnoreCase��True-�����ִ�Сд��False-���ִ�Сд��Ĭ�ϲ����ִ�Сд
'indexNum�����صڼ������ݣ�Ĭ�ϵ�һ�Ҳ����0��
Function v_regexContent(oriText, patternText, Optional isGlobal = True, Optional isIgnoreCase = True, Optional indexNum = 0)
    With CreateObject("VBscript.regexp")
        .Pattern = patternText
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        If .Execute(oriText).Count = 0 Then
            v_regexContent = ""
        Else
            v_regexContent = .Execute(oriText)(indexNum)
        End If
    End With
End Function

'�滻�ı������з���������ַ�������
'oriText��Ŀ���ַ�����patternText��������ʽ
'replaceText���滻�ɵ��ַ�����Ĭ��ȥ����Ҳ����""
'isGlobal��True-ƥ�����У�False-ƥ���һ�������Ĭ��ƥ������
'isIgnoreCase��True-�����ִ�Сд��False-���ִ�Сд��Ĭ�ϲ����ִ�Сд
Function v_regexReplace(oriText, patternText, Optional replaceText = "", Optional isGlobal = True, Optional isIgnoreCase = True)
    With CreateObject("VBscript.regexp")
        .Pattern = patternText
        .Global = isGlobal
        .IgnoreCase = isIgnoreCase
        replaceText = CStr(replaceText)
         v_regexReplace = .Replace(oriText, replaceText)
     End With
End Function


'#########################�����������##########################








