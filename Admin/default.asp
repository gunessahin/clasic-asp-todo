<!--
    Application: Clasic Asp Todo Admin v0.0.0.1
    https://gunessahin.github.io/
    ════════════════════════════════════════════════════════════════════════════════════════════════════
    Copyright gunessahin@outlook.com.tr
    Licence (https://github.com/gunessahin)
-->

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Clasic Asp Todo Sample | Standart</title>
</head>
<body>
    <%
    ' Windows için gerekli Acces Database Engine
    ' * [Database (Acces Database Engine)](https://www.microsoft.com/en-us/download/details.aspx?id=13255)
    ' Bir veri kaynağına bağlantı kurmak için kullanılan bilgileri gösterir.
    %>
    <!-- Bağlantı Oluşturulması -->
    <%
    ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/connectionstring-property-ado
    ' Bağlantı Terimi
    ' * Enable Parent Path On IIS https://docs.microsoft.com/en-us/iis/application-frameworks/running-classic-asp-applications-on-iis-7-and-iis-8/classic-asp-parent-paths-are-disabled-by-default
    ConnectionString                            =   "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/App_Data/test/data.mdb")
    
    ' Bağlantı Tanımlanıyor. 
    Set ActiveConnection                        =   Server.CreateObject("ADODB.Connection")
    
    ' Bağlantı Modu
    ' Bir Bağlantı , Kayıt veya Akış nesnesindeki verileri değiştirmeye yönelik mevcut izinleri belirtir.
    ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/mode-property-ado
    ActiveConnection.Mode                       =   0
                                                'adModeUnknown          0	        İzinler ayarlanmadı veya belirlenemedi.
                                                'adModeRead             1	        Salt Okunur.
                                                'adModeWrite            2	        Salt Yazma.
                                                'adModeReadWrite	    3	        Okuma / Yazma.
                                                'adModeShareDenyRead	4           Diğerlerinin okuma izinleriyle bir bağlantı açmasını engeller.
                                                'adModeShareDenyWrite	8	        Başkalarının yazma izinleriyle bağlantı açmasını engeller.
                                                'adModeShareExclusive	12	        Başkalarının bağlantı kurmasını engeller.
                                                'adModeShareDenyNone	16	        Başkalarının herhangi bir izinle bir bağlantı kurmasına izin verir.
                                                'adModeRecursive        0x400000	Geçerli Kaydın tüm alt kayıtlarında izinleri ayarlamak için adModeShareDenyNone, adModeShareDenyWrite veya adModeShareDenyRead ile birlikte kullanılır.
        
    ' Bağlantının, bağlantı terimi tanımlanıyor.
    ActiveConnection.ConnectionString           =   ConnectionString
    
    ' Bağlantı açılıyor.
    ActiveConnection.Open
    %>

    <!-- T-SQL Çalıştırma -->
    <%
        ' Transact-SQL Yordamı
        tSQL                                    =   "SELECT * FROM Todo"

        ' Belirtilen sorguyu, SQL deyimini, saklı yordamı veya sağlayıcıya özgü metni yürütür
        ' Transact-SQL olarak sağlayıcaya iletir.
        ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/execute-method-ado-connection
        ' * Enable 32-bit applications on iis
        Set data                                =   ActiveConnection.Execute(tSQL)
    %>

    <!-- Tablo Verilerinde tarama -->
    <%
    ' Data Taramaya Başlıyor
    While Not data.EOF
    
        ' Data'dan Modele aktar
        Set itemName                            =   data("Name")
        Set itemID                              =   data("ID")
    
        ' Yazdırılıyor
        Response.Write itemID & itemName & "<br>"
    
        ' Aramaya devam ediyor. ( Kursörü bir sonraki kayda konumlandır )
        data.MoveNext
    
    ' Tüm kayıtlar tarandı.
    Wend
    %>

    <!-- Aktif bağlantının sonlandırılması -->
    <%
    ' Bağlantı kontrol ediliyor.
    if not ActiveConnection is nothing then
    
        ' Bağlantı Tanımlanıyor.
        ' İlişkili sistem kaynaklarını boşaltmak için Close yöntemini kullanarak bir Connection , Record , Recordset veya Stream nesnesini kapatın 
        ' Referans : https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/close-method-ado
        ActiveConnection.Close
    
    end if
    %>
</body>
</html>
