<!--#include virtual                                        =   "/Core/Data.asp"-->
<!--#include virtual                                        =   "/Core/Config.asp"-->
<%
    ' Database Hazırlanıyor.
    Set db                                                  =   New Database
    
    ' Bağlantı terimi, konfigurasyon dosyasından alınıyor.
    db.ConnectionString                                     =   config.DataConnectionString
%>