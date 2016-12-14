<%
class Underscore

    public property get VERSION
        VERSION = "0.1.0"
    end property

    private sub class_initialize()
    end sub

    private sub class_terminate()
    end sub

    public function map(byval collection, byref iteratee)
        dim index
        dim length : length = ubound(collection)
        dim func : set func = getref(iteratee)
        dim results() : redim results(length)
        for index = 0 to length
            dim item : set item = collection(index)
            results(index) = func(item, index)
            set item = nothing
        next
        map = results
        set func = nothing
    end function

    public sub forEach(byval collection, byref iteratee)
        dim index
        dim length : length = ubound(collection)
        dim func : set func = getref(iteratee)
        for index = 0 to length
            dim item : set item = collection(index)
            func item, index
            set item = nothing
        next
        set func = nothing
    end sub

    public function filter(byval collection, byref predicate)
        dim index
        dim func : set func = getref(predicate)
        dim results() : redim results(-1)
        for index = 0 to ubound(collection) - 1
            dim item : set item = collection(index)
            dim result : result = func(item, index)
            if (result) then
                redim preserve results (ubound(results) + 1)
                set results(ubound(results)) = item
            end if
            set item = nothing
        next
        filter = results
        set func = nothing
    end function

end class
%>