' от Комарова Ильи
' компиляция файла с включением DLL в EXE файл


    Public Sub New()
        'см. описание в коде функции CurrentDomain_AssemblyResolve
        'данная строка должна идти до InitializeComponent
#If CONFIG = "Release" Then
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve
#End If
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавьте все инициализирующие действия после вызова InitializeComponent().

    End Sub
    ''' <summary>
    ''' Функция-обработчик события AppDomain.CurrentDomain.AssemblyResolve (коду требуется библиотека, но он не может
    ''' ее найти). В этом случае извлекаем библиотеку из ресурсов и предоставляем коду.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    ''' <returns></returns>
    ''' <remarks>http://lindsaybradford.wordpress.com/2012/08/29/packing-vb-net-assemblies-into-a -single-executable/</remarks>
    Private Shared Function CurrentDomain_AssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
        Dim resourceName As String = "MeasurementsChart." + New AssemblyName(args.Name).Name + ".dll"
        'получаем ресурс из сборки по его имени, тут же помещая в созданный поток
        Using stream = Assembly.GetExecutingAssembly.GetManifestResourceStream(resourceName)
            'Считываем ресурс в массив байтов
            Dim assemblyData = New Byte(stream.Length - 1) {}
            stream.Read(assemblyData, 0, assemblyData.Length)
            stream.Close()
            'Загружаем сборку из массива байтов в текущий домен приложения и возвращаем её
            Return Assembly.Load(assemblyData)
        End Using ' stream
    End Function