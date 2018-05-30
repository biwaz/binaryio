option explicit

dim objStream, objUTF8, objDOM, objBinHex
set objStream = createobject("ADODB.Stream")
objStream.charset = "Shift_JIS"

set objUTF8 = createobject("System.Text.UTF8Encoding")
' utf8->text objUTF8.GetString(bytes)
' text->utf8 objUTF8.GetBytes_4(strtext)

set objDOM = createobject("MSXML2.DOMDocument")

set objBinHex = objDOM.createelement("binhex")
objBinHex.dataType = "bin.hex"
' bin->hex objBinHex.nodetypedvalue = bytes : objBinHex.text
' hex->bin objBinHex.text = strtext : objBinHex.nodetypedvalue

public function readbinaryfile(filepath)
	with objStream
		.type = 1
		.open
		.loadfromfile filepath
		readbinaryfile = .read
		.close
	end with
end function

public sub savebinarydata(filepath, bytes)
	with objStream
		.type = 1
		.open
		.write bytes
		.savetofile filepath, 2
		.flush
		.close
	end with
end sub

function fromstring(encode, strtext)
	dim switch
	switch = vartype(encode)
	if switch < 2 then
		with objStream
			.open
			.type = 2
			.writetext strtext
			.position = 0
			.type = 1
			fromstring = .read
			.close
		end with
	elseif switch = 8 then
		with createobject("ADODB.Stream")
			.charset = encode
			.open
			.type = 2
			.writetext strtext
			.position = 0
			.type = 1
			fromstring = .read
			.close
		end with
	else
		with encode
			.open
			.type = 2
			.writetext strtext
			.position = 0
			.type = 1
			fromstring = .read
			.close
		end with
	end if
end function

function tostring(encode, bytes)
	dim switch
	switch = vartype(encode)
	if switch < 2 then
		with objStream
			.open
			.type = 1
			.write bytes
			.position = 0
			.type = 2
			tostring = .readtext
			.close
		end with
	elseif switch = 8 then
		with createobject("ADODB.Stream")
			.charset = encode
			.open
			.type = 1
			.write bytes
			.position = 0
			.type = 2
			tostring = .readtext
			.close
		end with
	else
		with encode
			.open
			.type = 1
			.write bytes
			.position = 0
			.type = 2
			tostring = .readtext
			.close
		end with
	end if
end function

class binaryio
	public buffer

	private sub class_initialize
		buffer = ""
	end sub

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

	public sub serialize_bytes(varname)
		objBinHex.nodetypedvalue = varname
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

	public function serialize(varname)
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
			serialize_bytes fromstring(null, varname)
		elseif typenum = 11 then
			serialize_bool varname
		elseif typenum = 17 then
			serialize_byte varname
		else
			serialize = false
		end if
	end function


	public function deserialize_bool
		deserialize_bool = not (left(buffer, 2) = "00")
		buffer = mid(buffer, 3)
	end function

	public function deserialize_byte
		deserialize_byte = cbyte("&h" & left(buffer, 2))
		buffer = mid(buffer, 3)
	end function

	public function deserialize_int
		deserialize_int = cint("&h" & mid(buffer, 3, 2) & left(buffer, 2))
		buffer = mid(buffer, 5)
	end function

	public function deserialize_lng
		deserialize_lng = clng("&h" & mid(buffer, 7, 2) & mid(buffer, 5, 2) & mid(buffer, 3, 2) & left(buffer, 2))
		buffer = mid(buffer, 9)
	end function

	public function deserialize_date
		dim num0, num1
		num0 = cdbl("&h" & mid(buffer, 15, 2) & mid(buffer, 13, 2) & mid(buffer, 11, 2) & mid(buffer, 9, 2) & mid(buffer, 7, 2) & mid(buffer, 5, 2) & mid(buffer, 3, 1))
		buffer = mid(buffer, 14)

		num1 = int(num0 / &hC92A69C)
		num0 = ((num0 - num1 * &hC92A69C) * 32 + cint("&h" & mid(buffer, 4, 1) & left(buffer, 1)) / 8) / 78125
		buffer = mid(buffer, 4)

		deserialize_date = dateadd("s", num0, dateadd("d", num1, "1960/01/01"))
	end function

	public function deserialize_bytes(n)
		if n < 0 then
			dim num
			num = 0
			deserialize_bytes = ""
			do while true
				num = instr(num + 1, buffer, "00")
				if num mod 2 = 1 then
					objBinHex.text = left(buffer, num + 1)
					deserialize_bytes = objBinHex.nodetypedvalue
					buffer = mid(buffer, num + 2)
					exit function
				end if
			loop
			objBinHex.text = buffer
			deserialize_bytes = objBinHex.nodetypedvalue
			buffer = ""
		else
			objBinHex.text = left(buffer, n * 2)
			deserialize_bytes = objBinHex.nodetypedvalue
			buffer = mid(buffer, n * 2 + 1)
		end if
	end function

	public function deserialize_double
		dim num, sign
		num = cint("&h" & mid(buffer, 15, 2) & mid(buffer, 13, 1))
		if num and &h800 then
			num = num - &hbff
			sign = -1
		else
			num = num - &h3ff
			sign = 1
		end if
		deserialize_double = sign * (cdbl("&h1" & mid(buffer, 14, 1) & mid(buffer, 11, 2) & mid(buffer, 9, 2) & mid(buffer, 7, 1)) + clng("&h" & mid(buffer, 8, 1) & mid(buffer, 5, 2) & mid(buffer, 3, 2) & left(buffer, 2)) / &h10000000) / &h1000000 * exp(0.693147180559945 * num)
		buffer = mid(buffer, 17)
	end function

	public function deserialize(varname)
		dim typenum
		typenum = vartype(varname)
		deserialize = true
		if     typenum = 2 then
			varname = deserialize_int
		elseif typenum = 3 then
			varname = deserialize_lng
		elseif typenum = 5 then
			varname = deserialize_double
		elseif typenum = 7 then
			varname = deserialize_date
		elseif typenum = 8 then
			varname = tostring(null, deserialize_bytes(-1))
		elseif typenum = 11 then
			varname = deserialize_bool
		elseif typenum = 17 then
			varname = deserialize_byte
		else
			deserialize = false
		end if
	end function
end class
