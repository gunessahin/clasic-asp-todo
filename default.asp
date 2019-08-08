<!--
    Application: Clasic Asp Todo v0.0.0.1
    https://gunessahin.github.io/
    ════════════════════════════════════════════════════════════════════════════════════════════════════
    Copyright gunessahin@outlook.com.tr
    Licence (https://github.com/gunessahin)
-->
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Clasic Asp Todo Sample</title>
</head>
<body>
    <!--#include virtual                                        =   "/Data/Todo.asp"-->
    <%
        ' Todo Data Context
        Dim TodoData        :   Set TodoData                    =   new TodoContext
        
        ' Todo Data Table
        Dim TodoDataTable   :   Set TodoDataTable               =   TodoData.List()
    %>
    <table>
        <tbody>
            <%If (TodoDataTable is Nothing) then%>
            <tr>
                <td>Record not found</td>
            </tr>
            <%Else%>

                <%Dim item  :   For Each item in TodoDataTable.Items %>
                <tr>
                    <td><%=item.ID %></td>
                    <td><%=item.Name %></td>
                </tr>
                <%Next%>

            <%End If%>
        </tbody>
    </table>

    <hr>

    <!-- CRUID İşlemleri -->
    <%
        ' Veri Bulunuyor
        Set todoItem = TodoData.Find(1)
        Response.Write "Bulundu : " & todoItem.Name & "<br>"
    
        ' Veri Ekleniyor
        ' TodoData.Add "Active Server Page - Data Insert : " & Now()
        ' Response.Write "Eklendi <br>"

        ' Veri Güncelleniyor
        ' TodoData.Update 1, "Active Server Page - Data Update : " & Now()
        ' Response.Write "Güncellendi <br>"

        ' Veri Siliniyor
        ' TodoData.Delete 2
        ' Response.Write "Silindi <br>"
    %>
</body>
</html>
