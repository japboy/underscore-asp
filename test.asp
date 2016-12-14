<%
Option Explicit
response.addheader "Content-type", "text/markdown; charset=UTF-8"
sub assert(cond1, cond2, mesg)
    dim text
    if (mesg = "") then
        mesg = "ok."
    end if
    if (cond1 = cond2) then
        text = "- [x] " & mesg & vbLf
    else
        text = "- [ ] " & mesg & vbLf
    end if
    response.write text
end sub
%>
<!--#include file="underscore.asp" -->

<%
dim U : set U = new Underscore
%>


# Version

<% assert "0.1.0", U.VERSION, "Expected version is 0.1.0." %> 


# `.map` Function

Return 10 list items multiplying by 2:

<%
dim idx1, collection1(10)
for idx1 = 0 to ubound(collection1)
    dim dictionary1 : set dictionary1 = server.createobject("Scripting.Dictionary")
    dictionary1.add "index", idx1
    set collection1(idx1) = dictionary1
next

function multiply2(dic, idx)
    multiply2 = dic.item("index") * 2
end function

dim results1 : results1 = U.map(collection1, "multiply2")
%>

<% assert 0, results1(0), "Expected value of index 0 is 0." %>
<% assert 2, results1(1), "Expected value of index 1 is 2." %>
<% assert 4, results1(2), "Expected value of index 2 is 4." %>
<% assert 6, results1(3), "Expected value of index 3 is 6." %>
<% assert 8, results1(4), "Expected value of index 4 is 8." %>
<% assert 10, results1(5), "Expected value of index 5 is 10." %>
<% assert 12, results1(6), "Expected value of index 6 is 12." %>
<% assert 14, results1(7), "Expected value of index 7 is 14." %>
<% assert 16, results1(8), "Expected value of index 8 is 16." %>
<% assert 18, results1(9), "Expected value of index 9 is 18." %>


# `.forEach` Sub procedure

Return 10 list items multiplying by 3:

<%
dim idx2, collection2(10)
for idx2 = 0 to ubound(collection2)
    dim dictionary2 : set dictionary2 = server.createobject("Scripting.Dictionary")
    dictionary2.add "index", idx2
    set collection2(idx2) = dictionary2
next

dim results2 : redim results2(ubound(collection2))

sub multiply3(dic, idx)
    dim num : num = dic.item("index") * 3
    results2(idx) = num
end sub

U.forEach collection2, "multiply3"
%>

<% assert 0, results2(0), "Expected value of index 0 is 0." %>
<% assert 3, results2(1), "Expected value of index 1 is 3." %>
<% assert 6, results2(2), "Expected value of index 2 is 6." %>
<% assert 9, results2(3), "Expected value of index 3 is 9." %>
<% assert 12, results2(4), "Expected value of index 4 is 12." %>
<% assert 15, results2(5), "Expected value of index 5 is 15." %>
<% assert 18, results2(6), "Expected value of index 6 is 18." %>
<% assert 21, results2(7), "Expected value of index 7 is 21." %>
<% assert 24, results2(8), "Expected value of index 8 is 24." %>
<% assert 27, results2(9), "Expected value of index 9 is 27." %>


# `.filter` Function

Return 5 of 10 list items filtered by even numbers:

<%
dim idx3, collection3(10)
for idx3 = 0 to ubound(collection3)
    dim dictionary3 : set dictionary3 = server.createobject("Scripting.Dictionary")
    dictionary3.add "index", idx3
    set collection3(idx3) = dictionary3
next

function even(dic, idx)
    dim result : result = (dic.item("index") Mod 2 = 0)
    even = result
end function

dim results3 : results3 = U.filter(collection3, "even")
%>

<% assert 0, results3(0).item("index"), "Expected value of index 0 is 0." %>
<% assert 2, results3(1).item("index"), "Expected value of index 1 is 2." %>
<% assert 4, results3(2).item("index"), "Expected value of index 2 is 4." %>
<% assert 6, results3(3).item("index"), "Expected value of index 3 is 6." %>
<% assert 8, results3(4).item("index"), "Expected value of index 4 is 8." %>