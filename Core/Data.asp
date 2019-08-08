<!--
    Application: Clasic Asp Data Connection v0.0.0.1
    https://azmisahin.github.io/
    ════════════════════════════════════════════════════════════════════════════════════════════════════
    Copyright azmisahin@outlook.com
    Licence (https://github.com/azmisahin)
-->

<%
    ' Veritabanı Modeli.
    ' Veritabanına ve sağlayıcıyla nasıl ileteşim kurulacağını tanımlar.
    ' T-SQL sorgulama dili ile iletişim kurar.
    '================================================================================

    ' Database Sınıfı.
    class Database

        ' Veritabanı Access 200 ve üzeri için planlanmıştır.
        ' Sql Server Önerilen.
        public ActiveConnection 

        ' Bir veri kaynağına bağlantı kurmak için kullanılan bilgileri gösterir.
        ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/connectionstring-property-ado
        public ConnectionString             'Bağlantı terimi.
    
        ' Bir Bağlantı , Kayıt veya Akış nesnesindeki verileri değiştirmeye yönelik mevcut izinleri belirtir.
        ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/mode-property-ado
        private mode                        'Balantı Modu.
    
    
        ' Sınıf İlk Yapılandırması.
        public sub Class_Initialize()            
    
            ' İlk Yapılandırma.
            Set ActiveConnection                        =   Nothing
    
            ' Mode özelliği yalnızca Connection nesnesi kapatıldığında ayarlanabilir .
            ' Referans : https://www.w3schools.com/asp/prop_rec_mode.asp
            mode                                        =   3
    
                                'adModeUnknown          0	        İzinler ayarlanmadı veya belirlenemedi.
                                'adModeRead             1	        Salt Okunur.
                                'adModeWrite            2	        Salt Yazma.
                                'adModeReadWrite	    3	        Okuma / Yazma.
                                'adModeShareDenyRead	4           Diğerlerinin okuma izinleriyle bir bağlantı açmasını engeller.
                                'adModeShareDenyWrite	8	        Başkalarının yazma izinleriyle bağlantı açmasını engeller.
                                'adModeShareExclusive	12	        Başkalarının bağlantı kurmasını engeller.
                                'adModeShareDenyNone	16	        Başkalarının herhangi bir izinle bir bağlantı kurmasına izin verir.
                                'adModeRecursive        0x400000	Geçerli Kaydın tüm alt kayıtlarında izinleri ayarlamak için adModeShareDenyNone, adModeShareDenyWrite veya adModeShareDenyRead ile birlikte kullanılır.
        
        end sub

        ' Veritabanı Açılıyor.
        public function Open
    
            ' Bağlantı kontrol ediliyor.
            if ActiveConnection is nothing then
    
                ' Bağlantı Tanımlanıyor. 
                Set ActiveConnection                    =   CreateObject("ADODB.Connection")
    
                ' Bağlantı Modu
                ActiveConnection.Mode                   =   mode
    
                ' Bağlantının, bağlantı terimi tanımlanıyor.
                ActiveConnection.ConnectionString       =   ConnectionString
    
                ' Bağlantı açılıyor.
                ActiveConnection.Open
            end if
    
            ' Aktif bağlantı
            Set Open                                    =   ActiveConnection
    
        end function
    
        ' Veritabanı Kapatılıyor.
        public function Close
    
            ' Bağlantı kontrol ediliyor.
            if not ActiveConnection is nothing then
    
                ' Bağlantı Tanımlanıyor.
                ' İlişkili sistem kaynaklarını boşaltmak için Close yöntemini kullanarak bir Connection , Record , Recordset veya Stream nesnesini kapatın 
                ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/close-method-ado
                ActiveConnection.Close
    
                ' Balantı Nothing
                ' Bir nesneyi bellekten tamamen ortadan kaldırmak için nesneyi kapatın ve nesne değişkenini Nothing olarak ayarlayın.
                ActiveConnection                        =   nothing
    
            end if
    
            ' Aktif bağlantı
            Set Close                                   =   ActiveConnection
    
        end function

        ' Transact-SQL (T-SQL) kullanılan, SQL sorgulama dilindeki komut dizisini çalıştırır.
        public function Run(tSQL)
            
            On Error Resume Next                        ' Herhangi bir hata olması durumunda devam et
            '--------------------------------------------------------------------------------

            ' Bağlantı açılarak, tanımlanıyor.
            Set ActiveConnection                        =   Open
    
            ' Belirtilen sorguyu, SQL deyimini, saklı yordamı veya sağlayıcıya özgü metni yürütür
            ' Transact-SQL olarak sağlayıcaya iletir.
            ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/execute-method-ado-connection
            Set Run                                     =   ActiveConnection.Execute(tSQL)
            
            '--------------------------------------------------------------------------------
            ' Hata Yakalama
            If Err.Number                               <>  0 Then
                ' Bir Hata Oluştu
                
                ' Hata String Verisi
                connectionErrorString                   =   ""
                
                ' Aktif Bağlantıda hatalar aranıyor.
                for each connectionError in ActiveConnection.Errors
                    
                    ' Aktif bağlantıdaki hata 
                    connectionErrorString               =   connectionErrorString   &   connectionError 
                next
                
                ' Hata Fırlatmaya hazırlanıyor.
                exeption = Err.number & " | " & Err.Description & " | " & connectionErrorString
                
                ' Hata Fırlatılıyor
                Throw  exeption
                
            else
                ' Loglama Hazırlanıyor
                log = tSql

                ' Log Gönderiliyor
                Trace log

            End If

        end function
    
        ' Query İşlemleri
        public function Query(tSQL)

            ' Sorgu çalıştırılıyor.
            Set Query                                   =   Run(tSQL)
            
        end function
    
        ' Hata Fırlatma
        private function Throw(message)
            Response.Write("ERR:>")
            Response.Write(message)

        end function

        ' Log Fırlatma
        private function Trace(message)
            Response.Write("LOG:>")
            Response.Write(message)

        end function
    
    ' Sınıf Sonu
    end class
%>