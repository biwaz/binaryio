option explicit

dim objStream, objBinHex

set objStream = createobject("ADODB.Stream")
objStream.charset = "Shift_JIS"

set objBinHex = createobject("MSXML2.DOMDocument").createelement("binhex")
objBinHex.dataType = "bin.hex"
' bin2hex objBinHex.nodetypedvalue = bytes : objBinHex.text
' hex2bin objBinHex.text = strtext : objBinHex.nodetypedvalue

class binaryio
	public buffer

	private sub class_initialize
		buffer = ""
	end sub

	public sub readbinaryfile(filepath)
		objStream.type = 1
		objStream.open
		objStream.loadfromfile filepath
		objBinHex.nodetypedvalue = objStream.read
		objStream.close
		buffer = objBinHex.text
	end sub

	public sub savebinarydata(filepath)
		objStream.type = 1
		objStream.open
		objBinHex.text = buffer
		objStream.write objBinHex.nodetypedvalue
		objStream.savetofile filepath, 2
		objStream.flush
		objStream.close
	end sub

	public function bytes2string(bytes)
		objStream.open
		objStream.type = 1
		objStream.write bytes
		objStream.position = 0
		objStream.type = 2
		bytes2string = objStream.readtext
		objStream.close
	end function

	public function string2bytes(strtext)
		objStream.open
		objStream.type = 2
		objStream.writetext strtext
		objStream.position = 0
		objStream.type = 1
		string2bytes = objStream.read
		objStream.close
	end function


	public sub serialize_bool(varname)
		if varname then buffer = buffer & "FF" else buffer = buffer & "00"
	end sub

	public sub serialize_byte(varname)
		buffer = buffer & right("0" & hex(varname), 2)
	end sub

	public sub serialize_int(varname)
		dim s
		s = right("000" & hex(varname), 4)
		buffer = buffer & mid(s, 3, 2) & left(s, 2)
	end sub

	public sub serialize_lng(varname)
		dim s
		s = right("0000000" & hex(varname), 8)
		buffer = buffer & mid(s, 7, 2) & mid(s, 5, 2) & mid(s, 3, 2) & left(s, 2)
	end sub

	public sub serialize_date(varname)
		dim num0, num1, num2, s
		num0 = datediff("s", "1960/01/01", varname)
		num1 = num0 * &h12d
		num2 = int(num1 / &h1000)

		num0 = num0 * &h13 + num2
		num1 = num1 - num2 * &h1000
		num2 = int(num0 / &h2000)

		s = right("0000000" & hex(num2), 8) & right("000000" & hex(((num0 - num2 * &h2000) * &h1000 + num1) * 8), 7) & "0"
		buffer = buffer & mid(s, 15, 2) & mid(s, 13, 2) & mid(s, 11, 2) & mid(s, 9, 2) & mid(s, 7, 2) & mid(s, 5, 2) & mid(s, 3, 2) & left(s, 2)
	end sub

	public sub serialize_string(varname)
		objBinHex.nodetypedvalue = string2bytes(varname)
		buffer = buffer & objBinHex.text & "00"
	end sub

	public sub serialize_double(varname)
		dim num0, num1, num2, s
		if 0 = sgn(varname) then
			buffer = buffer & "0000000000000000"
		else
			if 0 < sgn(varname) then
				num0 = varname
				num1 = &h3ff
			elseif sgn(varname) < 0 then
				num0 = -varname
				num1 = &hbff
			end if

			num2 = int(log(num0)/log(2))
			s = hex(num2 + num1)
			num0 = int(exp(0.693147180559945 * (24 - num2)) + 0.5) * num0
			num1 = int(num0)

			s = s & mid(hex(num1), 2) & right("0000000" & hex(int((num0 - num1) * &h10000000)), 7)
			buffer = buffer & mid(s, 15, 2) & mid(s, 13, 2) & mid(s, 11, 2) & mid(s, 9, 2) & mid(s, 7, 2) & mid(s, 5, 2) & mid(s, 3, 2) & left(s, 2)
		end if
	end sub

	public sub serialize(varname)
		dim typenum
		typenum = vartype(varname)
		serialize = true
		if     typenum = 2 then
			serialize_int varname
		elseif typenum = 3 then
			serialize_lng varname
		elseif typenum = 5 then
			serialize_double varname
		elseif typenum = 7 then
			serialize_date varname
		elseif typenum = 8 then
			serialize_string varname
		elseif typenum = 11 then
			serialize_bool varname
		elseif typenum = 17 then
			serialize_byte varname
		else
			serialize = false
		end if
	end sub


	public sub deserialize_bool(varname)
		varname = not (left(buffer, 2) = "00")
		buffer = mid(buffer, 3)
	end sub

	public sub deserialize_byte(varname)
		varname = cbyte("&h" & left(buffer, 2))
		buffer = mid(buffer, 3)
	end sub

	public sub deserialize_int(varname)
		varname = cint("&h" & mid(buffer, 3, 2) & left(buffer, 2))
		buffer = mid(buffer, 5)
	end sub

	public sub deserialize_lng(varname)
		varname = clng("&h" & mid(buffer, 7, 2) & mid(buffer, 5, 2) & mid(buffer, 3, 2) & left(buffer, 2))
		buffer = mid(buffer, 9)
	end sub

	public sub deserialize_date(varname)
		dim num0, num1
		num0 = cdbl("&h" & mid(buffer, 15, 2) & mid(buffer, 13, 2) & mid(buffer, 11, 2) & mid(buffer, 9, 2) & mid(buffer, 7, 2) & mid(buffer, 5, 2) & mid(buffer, 3, 1))
		buffer = mid(buffer, 14)

		num1 = int(num0 / &hC92A69C)
		num0 = ((num0 - num1 * &hC92A69C) * 32 + cint("&h" & mid(buffer, 4, 1) & left(buffer, 1)) / 8) / 78125
		buffer = mid(buffer, 4)

		varname = dateadd("s", num0, dateadd("d", num1, "1960/01/01"))
	end sub

	public sub deserialize_string(varname)
		dim num
		num = 0
		varname = ""
		do while true
			num = instr(num + 1, buffer, "00")
			if num mod 2 = 1 then
				objBinHex.text = left(buffer, num - 1)
				varname = bytes2string(objBinHex.nodetypedvalue)
				buffer = mid(buffer, num + 2)
				exit sub
			end if
		loop
		objBinHex.text = buffer
		varname = bytes2string(objBinHex.nodetypedvalue)
		buffer = ""
	end sub

	public sub deserialize_double(varname)
		dim num, sign
		num = cint("&h" & mid(buffer, 15, 2) & mid(buffer, 13, 1))
		if num and &h800 then
			num = num - &hbff
			sign = -1
		else
			num = num - &h3ff
			sign = 1
		end if
		varname = sign * (cdbl("&h1" & mid(buffer, 14, 1) & mid(buffer, 11, 2) & mid(buffer, 9, 2) & mid(buffer, 7, 1)) + clng("&h" & mid(buffer, 8, 1) & mid(buffer, 5, 2) & mid(buffer, 3, 2) & left(buffer, 2)) / &h10000000) / &h1000000 * exp(0.693147180559945 * num)
		buffer = mid(buffer, 17)
	end sub

	public function deserialize(varname)
		dim typenum
		typenum = vartype(varname)
		deserialize = true
		if     typenum = 2 then
			deserialize_int varname
		elseif typenum = 3 then
			deserialize_lng varname
		elseif typenum = 5 then
			deserialize_double varname
		elseif typenum = 7 then
			deserialize_date varname
		elseif typenum = 8 then
			deserialize_string varname
		elseif typenum = 11 then
			deserialize_bool varname
		elseif typenum = 17 then
			deserialize_byte varname
		else
			deserialize = false
		end if
	end function
end class
