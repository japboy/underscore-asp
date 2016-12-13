<%
Option Explicit
response.addheader "Content-type", "text/plain; charset=UTF-8"
%>
<!--#include file="underscore.asp" -->

<%
dim U : set U = new Underscore
%>


# Version

Expected: 0.1.0
Actual: <%= U.VERSION %> 


# .map Function

<%
dim idx1, collection1(10)
for idx1 = 0 to ubound(collection1)
    dim dictionary1 : set dictionary1 = server.createobject("Scripting.Dictionary")
    dictionary1.add "index", idx1
    set collection1(idx1) = dictionary1
next

function multiply10(dic, idx)
    multiply10 = dic.item("index") * 10
end function

dim results : results = U.map(collection1, "multiply10")
%>

<%= True = (results(0) = 0) %>,
<%= True = (results(1) = 10) %>,
<%= True = (results(2) = 20) %>,
<%= True = (results(3) = 30) %>,
<%= True = (results(4) = 40) %>,
<%= True = (results(5) = 50) %>,
<%= True = (results(6) = 60) %>,
<%= True = (results(7) = 70) %>,
<%= True = (results(8) = 80) %>,
<%= True = (results(9) = 90) %>,


# .forEach Function

<%
dim idx2, collection2(10)
for idx2 = 0 to ubound(collection2)
    dim dictionary2 : set dictionary2 = server.createobject("Scripting.Dictionary")
    dictionary2.add "index", idx2
    set collection2(idx2) = dictionary2
next

sub multiply8(dic, idx)
    dim result : result = dic.item("index") * 8
    response.write result & ", "
end sub

U.forEach collection2, "multiply8"
%>
