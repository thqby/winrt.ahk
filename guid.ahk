class GUID {
    __new(sguid:=unset) {
        this.Ptr := DllCall("msvcrt\malloc", "ptr", 16, "cdecl ptr")
        if IsSet(sguid)
            DllCall("ole32.dll\IIDFromString", "wstr", sguid, "ptr", this, "hresult")
        else
            NumPut("int64", 0, "int64", 0, this)
    }
    
    __delete() => DllCall("msvcrt\free", "ptr", this, "cdecl")

    CopyToPtr(ptr) {
        DllCall('msvcrt\memcpy', 'ptr', ptr, 'ptr', this, 'ptr', 16, 'cdecl')
    }
    static FromOffset(buf, offset) {
        return {ptr: buf.ptr + offset, _outer_: buf, __delete: _ => 0, base: this.Prototype}
    }
    
    static __new() {
        this.Prototype.DefineProp 'ToString', {call: GuidToString}
    }
    
    static prototype.Size := 16
}

GuidToString(guid) {
    buf := Buffer(78)
    DllCall("ole32.dll\StringFromGUID2", "ptr", guid, "ptr", buf, "int", 39)
    return StrGet(buf, "UTF-16")
}