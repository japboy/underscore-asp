<%
class Underscore

    public property get VERSION
        VERSION = "0.1.2"
    end property

    private sub class_initialize()
    end sub

    private sub class_terminate()
    end sub

    ' map
    public function map(byval collection, byref iteratee)
        if (typename(getref(iteratee)) <> "Object") then
            map = empty
            exit function
        end if
        select case typename(collection)
            ' JSONarray
            case "JSONarray"
                set map = mapForJsonarray(collection, iteratee)
            ' Array
            case "Variant()"
                map = mapForArray(collection, iteratee)
            case else
                map = empty
        end select
    end function

    private function mapForJsonarray(byval collection, byref iteratee)
        dim func, index, item, result
        dim jsonarr : set jsonarr = new JSONarray
        set func = getref(iteratee)
        for index = 0 to ubound(collection.items)
            set item = collection.itemat(index)
            result = func(item, index)
            jsonarr.push result
            set item = nothing
        next
        set func = nothing
        set mapForJsonarray = jsonarr
        set jsonarr = nothing
    end function


    private function mapForArray(byval collection, byref iteratee)
        dim func, index, item, length
        length = ubound(collection)
        dim arr() : redim arr(length)
        set func = getref(iteratee)
        for index = 0 to length
            set item = collection(index)
            arr(index) = func(item, index)
            set item = nothing
        next
        set func = nothing
        mapForArray = arr
    end function

    ' forEach
    public sub forEach(byval collection, byref iteratee)
        if (typename(getref(iteratee)) <> "Object") then
            exit sub
        end if
        select case typename(collection)
            ' JSONarray
            case "JSONarray"
                forEachForJsonarray collection, iteratee
            ' Array
            case "Variant()"
                forEachForArray collection, iteratee
        end select
    end sub

    private sub forEachForJsonarray(byval collection, byref iteratee)
        dim func, index, item
        set func = getref(iteratee)
        for index = 0 to ubound(collection.items)
            set item = collection.itemat(index)
            func item, index
            set item = nothing
        next
        set func = nothing
    end sub

    private sub forEachForArray(byval collection, byref iteratee)
        dim func, index, item
        set func = getref(iteratee)
        for index = 0 to ubound(collection)
            set item = collection(index)
            func item, index
            set item = nothing
        next
        set func = nothing
    end sub

    ' filter
    public function filter(byval collection, byref predicate)
        if (typename(getref(predicate)) <> "Object") then
            filter = empty
            exit function
        end if
        select case typename(collection)
            ' JSONarray
            case "JSONarray"
                set filter = filterForJsonarray(collection, predicate)
            ' Array
            case "Variant()"
                filter = filterForArray(collection, predicate)
            ' Fallback
            case else
                filter = empty
        end select
    end function

    private function filterForJsonarray(byval collection, byref predicate)
        dim func, index, item, result
        dim jsonarr : set jsonarr = new JSONarray
        set func = getref(predicate)
        for index = 0 to ubound(collection.items)
            set item = collection.itemat(index)
            result = func(item, index)
            if (result) then
                jsonarr.push item
            end if
            set item = nothing
        next
        set func = nothing
        set filterForArray = jsonarr
        set jsonarr = nothing
    end function

    private function filterForArray(byval collection, byref predicate)
        dim func, index, item, result
        dim arr() : redim arr(-1)
        set func = getref(predicate)
        for index = 0 to ubound(collection)
            set item = collection(index)
            result = func(item, index)
            if (result) then
                redim preserve arr(ubound(arr) + 1)
                set arr(ubound(arr)) = item
            end if
            set item = nothing
        next
        set func = nothing
        filterForArray = arr
    end function

    ' where
    public function where(byval collection, byref attributes)
        if (typename(attributes) <> "Dictionary") then
            where = empty
            exit function
        end if
        select case typename(collection)
            ' JSONarray
            case "JSONarray"
                set where = whereForJsonarray(collection, attributes)
            ' Array
            case "Variant()"
                where = whereForArray(collection, attributes)
            ' Fallback
            case else
                where = empty
        end select
    end function

    private function whereForJsonarray(byval collection, byref attributes)
        dim index, item, sum, key
        dim jsonarr : set jsonarr = new JSONarray
        for index = 0 to ubound(collection.items)
            set item = collection.itemat(index)
            sum = 0
            for each key in attributes.keys
                if (attributes.item(key) = item(key)) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(attributes.keys) + 1) then
                jsonarr.push item
            end if
            set item = nothing
        next
        set whereForJsonarray = jsonarr
        set jsonarr = nothing
    end function

    private function whereForArray(byval collection, byref attributes)
        dim index, item, sum, key
        dim arr() : redim arr(-1)
        for index = 0 to ubound(collection)
            set item = collection(index)
            sum = 0
            for each key in attributes.keys
                if (attributes.item(key) = item(key)) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(attributes.keys) + 1) then
                redim preserve arr(ubound(arr) + 1)
                set arr(ubound(arr)) = item
            end if
            set item = nothing
        next
        whereForArray = arr
    end function

end class
%>
