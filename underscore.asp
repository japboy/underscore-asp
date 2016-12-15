<%
class Underscore

    public property get VERSION
        VERSION = "0.1.3"
    end property

    private sub class_initialize()
    end sub

    private sub class_terminate()
    end sub

    ' truthy
    public function truthy(byval value)
        truthy = not not value
    end function

    ' falsy
    public function falsy(byval value)
        falsy = not value
    end function

    ' forEach
    public sub forEach(byval list, byref iteratee)
        if typename(getref(iteratee)) <> "Object" then
            exit sub
        end if
        select case typename(list)
            ' JSONarray
            case "JSONarray"
                forEachForJsonarray list, iteratee
            ' Array
            case "Variant()"
                forEachForArray list, iteratee
        end select
    end sub

    private sub forEachForJsonarray(byval list, byref iteratee)
        dim func, index, item
        set func = getref(iteratee)
        for index = 0 to ubound(list.items)
            func list.itemat(index), index
        next
        set func = nothing
    end sub

    private sub forEachForArray(byval list, byref iteratee)
        dim func, index, item
        set func = getref(iteratee)
        for index = 0 to ubound(list)
            func list(index), index
        next
        set func = nothing
    end sub

    ' map
    public function map(byval list, byref iteratee)
        if typename(getref(iteratee)) <> "Object" then
            map = empty
            exit function
        end if
        select case typename(list)
            ' JSONarray
            case "JSONarray"
                set map = mapForJsonarray(list, iteratee)
            ' Array
            case "Variant()"
                map = mapForArray(list, iteratee)
            case else
                map = empty
        end select
    end function

    private function mapForJsonarray(byval list, byref iteratee)
        dim func, index, item, result
        dim jsonarr : set jsonarr = new JSONarray
        set func = getref(iteratee)
        for index = 0 to ubound(list.items)
            result = func(list.itemat(index), index)
            jsonarr.push result
        next
        set func = nothing
        set mapForJsonarray = jsonarr
        set jsonarr = nothing
    end function

    private function mapForArray(byval list, byref iteratee)
        dim func, index, item, length
        length = ubound(list)
        dim arr() : redim arr(length)
        set func = getref(iteratee)
        for index = 0 to length
            arr(index) = func(list(index), index)
        next
        set func = nothing
        mapForArray = arr
    end function

    ' reduce
    public function reduce(list, iteratee, memo)
    end function

    ' find
    public function find(byval list, byref predicate)
    end function

    ' filter
    public function filter(byval list, byref predicate)
        if typename(getref(predicate)) <> "Object" then
            filter = empty
            exit function
        end if
        select case typename(list)
            ' JSONarray
            case "JSONarray"
                set filter = filterForJsonarray(list, predicate)
            ' Array
            case "Variant()"
                filter = filterForArray(list, predicate)
            ' Fallback
            case else
                filter = empty
        end select
    end function

    private function filterForJsonarray(byval list, byref predicate)
        dim func, index, item, result
        dim jsonarr : set jsonarr = new JSONarray
        set func = getref(predicate)
        for index = 0 to ubound(list.items)
            result = func(list.itemat(index), index)
            if truthy(result) then
                jsonarr.push list.itemat(index)
            end if
        next
        set func = nothing
        set filterForJsonarray = jsonarr
        set jsonarr = nothing
    end function

    private function filterForArray(byval list, byref predicate)
        dim func, index, item, result
        dim arr() : redim arr(-1)
        set func = getref(predicate)
        for index = 0 to ubound(list)
            result = func(list(index), index)
            if truthy(result) then
                redim preserve arr(ubound(arr) + 1)
                if isobject(list(index)) then
                    set arr(ubound(arr)) = list(index)
                else
                    arr(ubound(arr)) = list(index)
                end if
            end if
        next
        set func = nothing
        filterForArray = arr
    end function

    ' where
    public function where(byval list, byref properties)
        if typename(properties) <> "Dictionary" then
            where = empty
            exit function
        end if
        select case typename(list)
            ' JSONarray
            case "JSONarray"
                set where = whereForJsonarray(list, properties)
            ' Array
            case "Variant()"
                where = whereForArray(list, properties)
            ' Fallback
            case else
                where = empty
        end select
    end function

    private function whereForJsonarray(byval list, byref properties)
        dim index, item, sum, key
        dim jsonarr : set jsonarr = new JSONarray
        for index = 0 to ubound(list.items)
            sum = 0
            for each key in properties.keys
                if properties.item(key) = list.itemat(index)(key) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(properties.keys) + 1) then
                jsonarr.push list.itemat(index)
            end if
        next
        set whereForJsonarray = jsonarr
        set jsonarr = nothing
    end function

    private function whereForArray(byval list, byref properties)
        dim index, item, sum, key
        dim arr() : redim arr(-1)
        for index = 0 to ubound(list)
            sum = 0
            for each key in properties.keys
                if properties.item(key) = list(index)(key) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(properties.keys) + 1) then
                redim preserve arr(ubound(arr) + 1)
                if isobject(list(index)) then
                    set arr(ubound(arr)) = list(index)
                else
                    arr(ubound(arr)) = list(index)
                end if
            end if
        next
        whereForArray = arr
    end function

    ' findWhere
    public function findWhere(byval list, byref properties)
        if typename(properties) <> "Dictionary" then
            where = empty
            exit function
        end if
        select case typename(list)
            ' JSONarray
            case "JSONarray"
                set findWhere = findWhereForJsonarray(list, properties)
            ' Array
            case "Variant()"
                set findWhere = findWhereForArray(list, properties)
            ' Fallback
            case else
                findWhere = empty
        end select
    end function

    private function findWhereForJsonarray(byval list, byref properties)
        dim index, item, sum, key
        for index = 0 to ubound(list.items)
            sum = 0
            for each key in properties.keys
                if properties.item(key) = list.itemat(index)(key) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(properties.keys) + 1) then
                if isobject(list.itemat(index)) then
                    set findWhereForJsonarray = list.itemat(index)
                else
                    findWhereForJsonarray = list.itemat(index)
                end if
                exit function
            end if
        next
        findWhereForJsonarray = empty
    end function

    private function findWhereForArray(byval list, byref properties)
        dim index, item, sum, key
        for index = 0 to ubound(list)
            sum = 0
            for each key in properties.keys
                if properties.item(key) = list(index)(key) then
                    sum = sum + 1
                end if
            next
            if (sum > 0) and (sum = ubound(properties.keys) + 1) then
                if isobject(list(index)) then
                    set findWhereForArray = list(index)
                else
                    findWhereForArray = list(index)
                end if
                exit function
            end if
        next
        findWhereForArray = empty
    end function

    ' some
    public function some(byval list)
    end function

    ' contains
    public function contains(byval list, byref value)
    end function

    ' pluck
    public function pluck(byval list, byref propertyName)
        if typename(properties) <> "String" then
            pluck = empty
            exit function
        end if
        select case typename(list)
            ' JSONobject
            case "JSONobject"
                set pluck = pluckForJsonobject(list, propertyName)
            ' Scripting.Dictionary
            case "Dictionary"
                pluck = pluckForDictionary(list, propertyName)
            ' Fallback
            case else
                pluck = empty
        end select
    end function

    private function pluckForJsonobject(byval list, byref propertyName)
    end function

    private function pluckForDictionary(byval list, byref propertyName)
    end function

end class
%>
