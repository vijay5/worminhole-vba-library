' �� �������� ����
' ���������� ����� � ���������� DLL � EXE ����


    Public Sub New()
        '��. �������� � ���� ������� CurrentDomain_AssemblyResolve
        '������ ������ ������ ���� �� InitializeComponent
#If CONFIG = "Release" Then
        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve
#End If
        ' ���� ����� �������� ������������ ��� ������������.
        InitializeComponent()

        ' �������� ��� ���������������� �������� ����� ������ InitializeComponent().

    End Sub
    ''' <summary>
    ''' �������-���������� ������� AppDomain.CurrentDomain.AssemblyResolve (���� ��������� ����������, �� �� �� �����
    ''' �� �����). � ���� ������ ��������� ���������� �� �������� � ������������� ����.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="args"></param>
    ''' <returns></returns>
    ''' <remarks>http://lindsaybradford.wordpress.com/2012/08/29/packing-vb-net-assemblies-into-a -single-executable/</remarks>
    Private Shared Function CurrentDomain_AssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
        Dim resourceName As String = "MeasurementsChart." + New AssemblyName(args.Name).Name + ".dll"
        '�������� ������ �� ������ �� ��� �����, ��� �� ������� � ��������� �����
        Using stream = Assembly.GetExecutingAssembly.GetManifestResourceStream(resourceName)
            '��������� ������ � ������ ������
            Dim assemblyData = New Byte(stream.Length - 1) {}
            stream.Read(assemblyData, 0, assemblyData.Length)
            stream.Close()
            '��������� ������ �� ������� ������ � ������� ����� ���������� � ���������� �
            Return Assembly.Load(assemblyData)
        End Using ' stream
    End Function