<% Option Explicit %>
<!-- #include file="jsonObject.class.asp" -->
<!-- #include file="underscore.asp" -->

<% dim U : set U = new Underscore %>

<%
response.addheader "Content-type", "text/markdown; charset=UTF-8"

dim indexIn, jsonarrayIn, jsonobjectIn, arrayIn(), dictionaryIn, valueIn
dim indexOut, jsonarrayOut, jsonobjectOut, arrayOut, dictionaryOut, valueOut

sub assert(cond1, cond2, mesg)
    dim text
    if (cond1 = cond2) then
        text = "- [x] " & mesg & vbLf
    else
        text = "- [ ] " & mesg & vbLf
    end if
    response.write text
end sub
%>


# Version

<% assert "0.1.2", U.VERSION, "Expected version is 0.1.2." %>


# `.forEach` Sub

Return 10 items of JSONarray multiplied by 2:

<%
set jsonarrayIn = new JSONarray

for indexIn = 0 to 9
    set jsonobjectIn = new JSONobject
    jsonobjectIn.add "index", indexIn
    jsonarrayIn.push jsonobjectIn
    set jsonobjectIn = nothing
next

set jsonarrayOut = new JSONarray

sub multiply2(jsonobj, index)
    dim num : num = jsonobj("index") * 2
    jsonarrayOut.push num
end sub

U.forEach jsonarrayIn, "multiply2"

set jsonarrayIn = nothing

assert 10, ubound(jsonarrayOut.items) + 1, "Expected length of JSONarray is 10."
indexOut = 0
for each valueOut in jsonarrayOut.items
    assert indexOut * 2, valueOut, "Expected value of index " & indexOut & " is " & indexOut * 2 & "."
    indexOut = indexOut + 1
next

set jsonarrayOut = nothing
%>

Return 10 items of Array multiplied by 3:

<%
redim arrayIn(9)

for indexIn = 0 to ubound(arrayIn)
    set dictionaryIn = server.createobject("Scripting.Dictionary")
    dictionaryIn.add "index", indexIn
    set arrayIn(indexIn) = dictionaryIn
    set dictionaryIn = nothing
next

redim arrayOut(ubound(arrayIn))

sub multiply3(dict, index)
    dim num : num = dict.item("index") * 3
    arrayOut(index) = num
end sub

U.forEach arrayIn, "multiply3"

assert 10, ubound(arrayOut) + 1, "Expected length of Array is 10."
for indexOut = 0 to ubound(arrayOut)
    assert indexOut * 3, arrayOut(indexOut), "Expected value of index " & indexOut & " is " & indexOut * 3 & "."
next
%>


# `.map` Function

Return 10 items of JSONarray divided by 2:

<%
set jsonarrayIn = new JSONarray

for indexIn = 0 to 9
    set jsonobjectIn = new JSONobject
    jsonobjectIn.add "index", indexIn
    jsonarrayIn.push jsonobjectIn
    set jsonobjectIn = nothing
next

function divide2(jsonobj, index)
    divide2 = jsonobj("index") / 2
end function

set jsonarrayOut = U.map(jsonarrayIn, "divide2")

assert 10, ubound(jsonarrayOut.items) + 1, "Expected length of Array is 10."
indexOut = 0
for each valueOut in jsonarrayOut.items
    assert indexOut / 2, valueOut, "Expected value of index " & indexOut & " is " & indexOut / 2 & "."
    indexOut = indexOut + 1
next
%>


Return 10 items of Array divided by 2:

<%
redim arrayIn(9)

for indexIn = 0 to ubound(arrayIn)
    set dictionaryIn = server.createobject("Scripting.Dictionary")
    dictionaryIn.add "index", indexIn
    set arrayIn(indexIn) = dictionaryIn
next

function divide3(dict, index)
    divide3 = dict.item("index") / 3
end function

arrayOut = U.map(arrayIn, "divide3")

assert 10, ubound(arrayOut) + 1, "Expected length of Array is 10."
for indexOut = 0 to ubound(arrayOut)
    assert indexOut / 3, arrayOut(indexOut), "Expected value of index " & indexOut & " is " & indexOut / 3 & "."
next

erase arrayOut
%>


# `.filter` Function

Return 5 of 10 items of JSONarray filtered by odd numbers:

<%
set jsonarrayIn = new JSONarray

for indexIn = 0 to 9
    set jsonobjectIn = new JSONobject
    jsonobjectIn.add "index", indexIn
    jsonarrayIn.push jsonobjectIn
    set jsonobjectIn = nothing
next

function odd(jsonobj, index)
    odd = (jsonobj("index") mod 2 <> 0)
end function

set jsonarrayOut = U.filter(jsonarrayIn, "odd")

assert 5, ubound(jsonarrayOut.items) + 1, "Expected length of JSONarray is 5."
indexOut = 0
for each jsonobjectOut in jsonarrayOut.items
    assert indexOut * 2 + 1, jsonobjectOut("index"), "Expected value of index " & indexOut & " is " & indexOut * 2 + 1 & "."
    indexOut = indexOut + 1
next
%>


Return 5 of 10 items of Array filtered by even numbers:

<%
redim arrayIn(9)

for indexIn = 0 to ubound(arrayIn)
    set dictionaryIn = server.createobject("Scripting.Dictionary")
    dictionaryIn.add "index", indexIn
    set arrayIn(indexIn) = dictionaryIn
next

function even(dict, index)
    even = (dict.item("index") mod 2 = 0)
end function

arrayOut = U.filter(arrayIn, "even")

assert 5, ubound(arrayOut) + 1, "Expected length of Array is 5."
for indexOut = 0 to ubound(arrayOut)
    assert indexOut * 2, arrayOut(indexOut).item("index"), "Expected value of index " & indexOut & " is " & indexOut * 2 & "."
next

erase arrayOut
%>


# `.where` Function

Return 4 of 10 items of JSONarray filtered by mod 3:

<%
set jsonarrayIn = new JSONarray

for indexIn = 0 to 9
    set jsonobjectIn = new JSONobject
    jsonobjectIn.add "index", indexIn
    jsonobjectIn.add "value", indexIn mod 3
    jsonarrayIn.push jsonobjectIn
    set jsonobjectIn = nothing
next

set dictionaryIn = server.createobject("Scripting.Dictionary")
dictionaryIn.add "value", 0

set jsonarrayOut = U.where(jsonarrayIn, dictionaryIn)

set jsonarrayIn = nothing
set dictionaryIn = nothing

assert 4, ubound(jsonarrayOut.items) + 1, "Expected length of JSONarray is 4."
for indexOut = 0 to ubound(jsonarrayOut.items)
    assert 0, jsonarrayOut.itemat(indexOut)("value"), "Expected value of index " & indexOut & " is 0."
next
%>


Return 3 of 10 items of Array filtered by mod 4:

<%
redim arrayIn(9)

for indexIn = 0 to ubound(arrayIn)
    set dictionaryIn = server.createobject("Scripting.Dictionary")
    dictionaryIn.add "index", indexIn
    dictionaryIn.add "value", indexIn mod 4
    set arrayIn(indexIn) = dictionaryIn
    set dictionaryIn = nothing
next

set dictionaryIn = server.createobject("Scripting.Dictionary")
dictionaryIn.add "value", 0

arrayOut = U.where(arrayIn, dictionaryIn)

set dictionaryIn = nothing

assert 3, (ubound(arrayOut) + 1), "Expected length of Array is 3."
for indexOut = 0 to ubound(arrayOut)
    assert 0, arrayOut(indexOut)("value"), "Expected value of index " & indexOut & " is 0."
next

erase arrayOut
%>


# `.findWhere` Function

Return 1 of 10 items of JSONarray filtered by mod 1:

<%
set jsonarrayIn = new JSONarray

for indexIn = 0 to 9
    set jsonobjectIn = new JSONobject
    jsonobjectIn.add "index", indexIn
    jsonobjectIn.add "value", indexIn mod 1
    jsonarrayIn.push jsonobjectIn
    set jsonobjectIn = nothing
next

set dictionaryIn  = server.createobject("Scripting.Dictionary")
dictionaryIn.add "value", 0

set jsonobjectOut = U.findWhere(jsonarrayIn, dictionaryIn)

set jsonarrayIn = nothing
set dictionaryIn = nothing

assert "JSONobject", typename(jsonobjectOut), "Expected return is a JSONobject."
assert 0, jsonobjectOut("index"), "Expected index is 0."
assert 0, jsonobjectOut("value"), "Expected value is 0."

set jsonobjectOut = nothing
%>

Return 1 of 10 items of Array filtered by mod 2:

<%
redim arrayIn(9)

for indexIn = 0 to ubound(arrayIn)
    set dictionaryIn = server.createobject("Scripting.Dictionary")
    dictionaryIn.add "index", indexIn
    dictionaryIn.add "value", indexIn mod 2
    set arrayIn(indexIn) = dictionaryIn
next

set dictionaryIn = server.createobject("Scripting.Dictionary")
dictionaryIn.add "value", 0

set dictionaryOut = U.findWhere(arrayIn, dictionaryIn)

set dictionaryIn = nothing

assert "Dictionary", typename(dictionaryOut), "Expected return is a Scripting.Dictionary."
assert 0, dictionaryOut("index"), "Expected index is 0."
assert 0, dictionaryOut("value"), "Expected value is 0."

set dictionaryOut = nothing
%>
