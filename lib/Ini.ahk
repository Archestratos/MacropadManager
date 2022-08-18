
; Version: 2022.07.01.1
; Credit to /u/anonymous1184
; Usages and examples: https://redd.it/s1it4j
; https://gist.github.com/anonymous1184/737749e83ade98c84cf619aabf66b063

Ini(Path, Sync := true)
{
	return new Ini_File(Path, Sync)
}

#Include %A_LineFile%\..\Object.ahk

class Ini_File extends Object_Proxy
{

	;region Public

	Persist()
	{
		IniRead buffer, % this.__path
		sections := {}
		for _,name in StrSplit(buffer, "`n")
			sections[name] := true
		for name in this.__data {
			this[name].Persist()
			sections.Delete(name)
		}
		for name in sections
			IniDelete % this.__path, % name
	}

	Sync(Set := "")
	{
		if (!StrLen(Set))
			return this.__sync
		for name in this
			this[name].Sync(Set)
		return this.__sync := !!Set
	}
	;endregion

	;region Overload

	Delete(Name)
	{
		if (this.__sync)
			IniDelete % this.__path, % Name
	}
	;endregion

	;region Meta

	__New(Path, Sync)
	{
		ObjRawSet(this, "__path", Path)
		ObjRawSet(this, "__sync", false)
		IniRead buffer, % Path
		for _,name in StrSplit(buffer, "`n") {
			IniRead data, % Path, % name
			this[name] := new Ini_Section(Path, name, data)
		}
		this.Sync(Sync)
	}

	__Set(Key, Value)
	{
		isObj := IsObject(Value)
		base := isObj ? ObjGetBase(Value) : false
		if (isObj && !base)
		|| (base && base.__Class != "Ini_Section") {
			path := this.__path
			sync := this.__sync
			this[Key] := new Ini_Section(path, Key, Value, sync)
			return obj ; Stop, hammer time!
		}
	}
	;endregion

}

class Ini_Section extends Object_Proxy
{

	;region Public

	Persist()
	{
		IniRead buffer, % this.__path, % this.__name
		keys := {}
		for _,key in StrSplit(buffer, "`n") {
			key := StrSplit(key, "=")[1]
			keys[key] := true
		}
		for key,value in this {
			keys.Delete(key)
			value := StrLen(value) ? " " value : ""
			IniWrite % value, % this.__path, % this.__name, % key
		}
		for key in keys
			IniDelete % this.__path, % this.__name, % key
	}

	Sync(Set := "")
	{
		if (!StrLen(Set))
			return this.__sync
		return this.__sync := !!Set
	}
	;endregion

	;region  Overload

	Delete(Key)
	{
		if (this.__sync)
			IniDelete % this.__path, % this.__name, % key
	}
	;endregion

	;region  Meta

	__New(Path, Name, Data, Sync := false)
	{
		ObjRawSet(this, "__path", Path)
		ObjRawSet(this, "__name", Name)
		ObjRawSet(this, "__sync", Sync)
		if (!IsObject(Data))
			Ini_ToObject(Data)
		for key,value in Data
			this[key] := value
	}

	__Set(Key, Value)
	{
		if (this.__sync) {
			; Value := StrLen(Value) ? " " Value : ""
			IniWrite % Value, % this.__path, % this.__name, % key
		}
	}
	;endregion

}

;region Auxiliary

Ini_ToObject(ByRef Data)
{
	info := Data, Data := {}
	for _,pair in StrSplit(info, "`n") {
		pair := StrSplit(pair, "=",, 2)
		Data[pair[1]] := pair[2]
	}
}
;endregion
