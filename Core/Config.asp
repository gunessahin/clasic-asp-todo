<!--
    Application: Clasic Asp Application.Config v0.0.0.1
    https://azmisahin.github.io/
    ════════════════════════════════════════════════════════════════════════════════════════════════════
    Copyright azmisahin@outlook.com
    Licence (https://github.com/azmisahin)
-->

<%
    ' Konfigurasyon Tanımları
    class Configuration
    
        ' Uyglama Klasörü
        public AppData
        
        ' Veri Bağlantı Terimi
        public DataConnectionString

        ' Log Klasörü
        public LogFolder

        ' Hata Ayıklama Modu Durumu
        public Debug

        ' Hatalı Giriş Deneme Sınırı
        ' Hatalı giriş denemelerinde, hesap geçeci olarak kilitlenecektir.
        ' Kilitlenme sınırını belirtir.
        public AccessFailedCount

        ' Hatalı giriş denemeleri sonucu, Hatalı Giriş Deneme sınırı aşılır ise
        '   ne kadar süre ile kilitli kalacağı bilgisi verilmektedir.
        '   süre dakika cinsinden eklenir.
        public AccessFailedDuration


        ' Sınıf İlk Yapılandırması.
        public sub Class_Initialize()

            ' Hata Ayıylama Modu        '   Talep Yerel Bir Ortamdan Geldiğinde
            Debug                       =   true

            ' Hatalı Giriş Deneme Sınırı
            ' Hatalı giriş denemelerinde, hesap geçeci olarak kilitlenecektir.
            ' Kilitlenme sınırını belirtir.
            AccessFailedCount           =   5

            ' Hatalı giriş denemeleri sonucu, Hatalı Giriş Deneme sınırı aşılır ise
            '   ne kadar süre ile kilitli kalacağı bilgisi verilmektedir.
            '   süre dakika cinsinden eklenir.
            AccessFailedDuration        =   1
            
            ' Ilk Yapılandırma 
            Init()

        end sub

        ' Konfigurasyon İlk Yapılandırma Tanımları
        private sub Init

            ' MapPath yöntemi sunucuda karşılık gelen fiziksel dizini belirtilen göreli veya sanal yolunu eşler.
            ' Referans : https://msdn.microsoft.com/en-us/library/ms524632(v=vs.90).aspx
    
            ' Sağlayıcı
            ' Geliştiricilerin verileri yazılıma bağlamasına yardımcı olur
            ' Refarans : https://www.connectionstrings.com/microsoft-jet-ole-db-4-0/
            '   Sql Server                    [   New     ]
            '   Data Source Name              [   Mybe    ]
            '   Joint Engine Technology       [   Red     ]

            ' Yerel Hata Ayıklama Modunda mı çalışıyor?
            
            if Debug = true Then
                ' Test Veritabanında çalış
                AppData                 =   "test"
            
            else
                ' Canlı Ortamda Çalış
                AppData                 =   "app"
            end if
    
            ' Log Klasörü
            LogFolder                   =   "/App_Data/" & AppData & "/log/"            
       
     end sub
    
    end class
    ' Genel Tanımlar

    Dim config  :   Set config = New Configuration
%>

<%

    ' DATA Model İşlemleri İçin Bağlantı Terimi.
    '--------------------------------------------------         

	' Sql Server Yapılandırması
    'config.DataConnectionString        =   "Provider = SQLOLEDB; Data Source = .\SQLEXPRESS; Persist Security Info=True; User ID = sa; Password = 123456; Initial Catalog = " & config.AppData & "_" & "data;"

    ' ODBC Yapılandırması
    'config.DataConnectionString        =   "DSN = " & config.AppData & "_" & "data; Uid = sa; Pwd = 123456;"
            
    ' Access New Version
    'config.DataConnectionString       =   "Driver={Microsoft Access Driver (*.mdb, *.accdb)};ExtendedAnsiSQL=1; DBQ=" & Server.MapPath("/App_Data/" & config.AppData & "/data.mdb")

    ' Standart JET            
    ' * [Database (Acces Database Engine)](https://www.microsoft.com/en-us/download/details.aspx?id=13255)
    ' * Enable Parent Path On IIS https://docs.microsoft.com/en-us/iis/application-frameworks/running-classic-asp-applications-on-iis-7-and-iis-8/classic-asp-parent-paths-are-disabled-by-default
    config.DataConnectionString        =   "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/App_Data/" & config.AppData & "/data.mdb")
%>