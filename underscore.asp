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
        dim idx
        dim len : len = ubound(collection)
        dim fn : set fn = getref(iteratee)
        dim results() : redim results(len)
        for idx = 0 to len
            dim item : set item = collection(idx)
            results(idx) = fn(item, idx)
            set item = nothing
        next
        map = results
        set fn = nothing
    end function

    public sub forEach(byval collection, byref iteratee)
        dim idx
        dim len : len = ubound(collection)
        dim fn : set fn = getref(iteratee)
        for idx = 0 to len
            dim item : set item = collection(idx)
            fn item, idx
            set item = nothing
        next
        set fn = nothing
    end sub

end class
%>