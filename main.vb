Sub test()
    On Error Resume Next '屏蔽错误，遇到错误继续执行
    Dim wordapp As Object, mytab As Object
    Dim mydoc
    Dim dirpath$, docname$, myfile$
    '创建word应用程序对象
    Set wordapp = CreateObject("word.application")
    '定义word文件存放路径
    dirpath = ThisWorkbook.Path & "/word/"
    '获取路径下word文档合集
    myfile = Dir(dirpath & "*.doc*")
    i = 2
    '循环读取文件夹下文档
    Range("a2:h10000").ClearContents
    Do While myfile <> ""
        '打开文件
        Set mydoc = wordapp.Documents.Open(dirpath & myfile)
        '获取表格对象
        Set mytab = mydoc.tables(1)
        '获取客户资料所在行数
        For j=1 To 100 
            temp = Left(Trim(mytab.Cell(j,1).Range.Text), Len("客户资料"))
            If temp="客户资料" then
                k = j
                Exit For
            End If
        Next
        '获取word单元格内容，存储到excel
        Cells(i,1) = nn(tt(Trim(mytab.Cell(k+1,2).Range.Text)))
        Cells(i,2) = dd(tt(Trim(mytab.Cell(k+1,3).Range.Text)))
        Cells(i,3) = tt(Trim(mytab.Cell(k+2,2).Range.Text))
        Cells(i,4) = tt(Trim(mytab.Cell(k+3,2).Range.Text))
        myfile = Dir
        i = i+1
    Loop
    '关闭文档
    mydoc.Close False
    '退出word应用程序
    wordapp.Quit
End Sub

'自定义函数，用于截断获取内容的最后一个特殊字符
Function tt(rtext)
    tt = Left(rtext, Len(rtext)-1)
End Function

'自定义函数，用于删除＂收货人姓名：＂
Function nn(rtext)
    nn = Right(rtext, Len(rtext)-Len("收货人姓名："))
End Function

'自定义函数，用于删除＂收货人姓名：＂
Function dd(rtext)
    dd = Right(rtext, Len(rtext)-Len("联系电话："))
End Function
