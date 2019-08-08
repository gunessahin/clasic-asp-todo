<!--
    Application: Clasic Asp Todo Data Model v0.0.0.1
    https://gunessahin.github.io/
    ════════════════════════════════════════════════════════════════════════════════════════════════════
    Copyright gunessahin@outlook.com.tr
    Licence (https://github.com/gunessahin)
-->

<!--#include virtual                                        =   "/Core/init.asp"-->
<%
    ' Todo Sınıfı
    class Todo
    
        ' SINIF ÖZELLİKLERİ
        '================================================================================    

        public ID                           'int                        Benzersiz Numara
        public Name                         'string   50                Ad
    
        ' Veritabanı ile bağımlı.
        public function Map(data)

            ID                              =   data("ID")              'int                        Benzersiz Numara
            Name                            =   data("Name")            'string   50                Ad

        end function
   
    end class

    ' Todo Context Sınıfı
    class TodoContext

        ' Table Models
        Public Models
    
        ' Bağlantı açıldı.
        Public data
    
        ' Sınıf İlk Yapılandırması
        public sub Class_Initialize()

            ' Table Model
            Set Models                              =   Server.CreateObject("Scripting.Dictionary")

        end sub
    
        ' Datayı Modele çevirir
        Private Function fromData(data)
    
            ' Bir veya daha fazla değişken için depolama alanını bildirir ve atar.
            ' Referans : https://docs.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/dim-statement
            Dim item

            ' Bildiren bir Set bir özellik için bir değer atamak için kullanılan özellik yordam
            ' Referans : https://docs.microsoft.com/tr-tr/dotnet/visual-basic/language-reference/statements/set-statement
            Set item                                =   new Todo

            ' Geçerli kayıt konumu, bir Recordset nesnesinin son kaydından sonra gelip gelmediğini gösteren bir değeri döndürür . Salt okunur Boolean
            if data.EOF then

                ' Kayıt Yok
                Set item                            =   Nothing

            else
                
                ' Data Kayıtları model özelliklerini aktarılıyor
                item.Map(data)
        
            end if
            
            ' Fonksiyon Dönüşü
            Set fromData                            =   item

        End Function

        ' List
        Public Function List()

            ' T-SQL 
            query                                                   =   "SELECT * FROM Todo"

            ' Data Table
            set data                                                =   db.Query(query)

            ' Index
            i                                                       =   0

            ' Data Taramaya Başlıyor
            While Not data.EOF
    
                ' Data'dan Modele aktar
                Set item                                            =   fromData(data)

               ' Listeye modeli ekle
                Models.Add                                              i, item

                ' Index i arttır
                i                                                   =   i   +   1
    
                ' Aramaya devam ediyor. ( Kursörü bir sonraki kayda konumlandır )
                data.MoveNext
    
            ' Tüm kayıtlar tarandı.
            Wend

            ' Models items listesini geri donusu
            Set List                                                =   Models

        End Function

        ' Bul    
        Public Function Find(id)

            ' T-SQL 
            query                                                   =   "SELECT * FROM Todo WHERE ID = " & id

            ' Data Table
            set data                                                =   db.Query(query)

            ' 
            Set Find                                                =   fromData(data)

        End Function

        ' Yeni Oge Ekle
        Public Function Add(itemName)

            ' T-SQL 
            query                                                   =   "INSERT INTO Todo(Name) VALUES(" & "'" & itemName & "'" & ")"

            ' Data Table
            set data                                                =   db.Query(query)

            ' 
            Set Add                                                 =   data

        End Function
    
        ' Guncelle    
        Public Function Update(id, itemName)

            ' T-SQL 
            query                                                   =   "UPDATE Todo SET Name = '" & itemName & "' WHERE ID = " & id

            ' Data Table
            set data                                                =   db.Query(query)

            ' 
            Set Update                                              =   data

        End Function

        ' Sil    
        Public Function Delete(id)

            ' T-SQL 
            query                                                   =   "DELETE FROM Todo WHERE ID = " & id

            ' Data Table
            set data                                                =   db.Query(query)

            ' 
            Set Delete                                              =   data

        End Function

    end class
%>