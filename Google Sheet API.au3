Global $GGSheet_CryptoJS_FILE = ""
Global Const $GGSheet_CryptoJS_URL = "https://raw.githubusercontent.com/aipit2/jsrsasign/main/jsrsasign.js"
Global $GGSheet_CryptoJS = ""
Global $GGSheet_Access_Token = ""
Global $GGSheet_Service_Account_File = ""
Global $GGSheet_SpreadSheet_Id = ""
Global $GGSheet_Header = ""
; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Setup
; Description ...: Kiểm tra có thư viện không - Kiểm tra xem có file jsrsasign.js trong thư mục chính không, nếu không thì load thư viện trên github
; Syntax ........: GGSheet_Setup()
; Parameters ....: None
; Return values .: Success 	-  URL thư viện
; 				 : @error 	-  Không load được thư viện trong máy và cả trên github
; Author ........: Trần Hùng
; ===============================================================================================================================

Func GGSheet_Setup($service_account_file_dir, $library_file_dir, $spread_sheet_id)
	Local $result
	Local $fileOpen = FileOpen($GGSheet_CryptoJS_FILE)
	Local $library = FileRead($fileOpen)
	If @error Then
		$library = _HttpRequest(2,$GGSheet_CryptoJS_URL)
		If @error Or $library == "" Then
			Return SetError(1,0,"Khởi tạo UDF thất bại, vui lòng kiểm tra lại file thư viện và link Github")
		EndIf
		$GGSheet_CryptoJS = $GGSheet_CryptoJS_URL
	Else
		$GGSheet_CryptoJS = $library_file_dir
	EndIf
	FileClose($fileOpen)

	$GGSheet_Service_Account_File = $service_account_file_dir

	Local $access_token = GGSheet_Create_Access_Token()
	If @error Then
		Return SetError(2,0,$access_token)
	EndIf

	$GGSheet_SpreadSheet_Id = $spread_sheet_id
	Return $access_token
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Create_JWT
; Description ...: Tạo JWT
; Syntax ........: GGSheet_Create_JWT($service_account_file_dir_or_data)
; Parameters ....: $service_account_file_dir_or_data - Đường dẫn hoặc data của service account.
; Return values .: Success 	- JWT dưới dạng string
;                : @error 	- Tạo JWT thất bại
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_Create_JWT()
	Local $data = FileRead($GGSheet_Service_Account_File)
	Local $jsonKey = _HttpRequest_ParseJSON($data)
	Local $header = '{"alg" : "RS256", "typ" : "JWT"}'
	Local $timeStamp = _GetTimeStamp()
	Local $claimSet = '{"iss": "' & $jsonKey.client_email & '",'
	$claimSet &= '"scope": "https://www.googleapis.com/auth/spreadsheets",'
	$claimSet &= '"aud": "https://www.googleapis.com/oauth2/v4/token",'
	$claimSet &= '"exp": ' & $timeStamp + 3600 & ','
	$claimSet &= '"iat": ' & $timeStamp & '}'
	Local $privateKey = StringRegExpReplace($jsonKey.private_key,'\n','\\n')
	Local $codeJs = 'var jwt = KJUR.jws.JWS.sign(null, ' & $header & ', ' & $claimSet & ', "' & $privateKey & '");'
	Local $JWT = _JS_Execute($GGSheet_CryptoJS,$codeJs,"jwt",True)
	If $JWT = "" Then
		Return SetError(1,0,"Tạo JWT thất bại")
	Else
		Return $JWT
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Create_Access_Token
; Description ...: Tạo access token (Có thời gian sử dụng trong 60 phút)
; Syntax ........: GGSheet_Create_Access_Token($service_account_file_dir_or_data)
; Parameters ....: $service_account_file_dir_or_data - Đường dẫn hoặc data của service account.
; Return values .: Access Token
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_Create_Access_Token()
	Local $jwt = GGSheet_Create_JWT()
	If @error Then
		Return SetError(1,0,"Tạo JWT thất bại")
	EndIf
	Local $url = "https://oauth2.googleapis.com/token"
	Local $data = "grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=" & _URIEncode($jwt)
	Local $res = _HttpRequest(2,$url,$data)
	Local $accessToken = StringRegExp($res,'access_token":"(.*?)"',1)
	If @error Then
		Return SetError(2,0,$res)
	EndIf
	$GGSheet_Header = "Authorization: Bearer " & $accessToken[0]
	Return $accessToken[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Read
; Description ...: Đọc giữ liệu từ vùng
; Syntax ........: GGSheet_Read($range)
; Parameters ....: $range               - Vùng muốn đọc.
; Return values .: Success	- Return sẽ là mảng 2 chiều chưa dữ liệu đọc được
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; Example .......: GGSheet_Read("Sheet1!A1:B4")
; ===============================================================================================================================
Func GGSheet_Read($range)
	Local $url = "https://sheets.googleapis.com/v4/spreadsheets/" & $GGSheet_SpreadSheet_Id & "/values/" & $range
	Local $res = _HttpRequest(2,$url,'','','',$GGSheet_Header)
	If GGSheet_Is_Access_Token_Expired($res) = 1 Then ; Access Token đã hết hạn --> Tạo mới
		GGSheet_Create_Access_Token()
		Return GGSheet_Read($range)
	EndIf

	Local $json = _HttpRequest_ParseJSON($res)

	If $json == False Then
		Return SetError(1,0,$res)
	ElseIf $json.values == False Then
		Return SetError(2,0,$res)
	EndIf

	Local $result = _Make2Array($json.values.toStr())
	If UBound($result) = 1 And UBound($result,2) = 1 Then ; Nếu đọc từ 1 ô thì trả về chuỗi
		Return $result[0][0]
	Else
		Return $result
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Write
; Description ...: Ghi dữ liệu vào vùng
; Syntax ........: GGSheet_Write($range, $value)
; Parameters ....: $range	- Vùng muốn ghi.
;                  $value	- Mảng 2 chiều chứa giá trị cần ghi - Nếu ghi giá trị vào 1 ô thì $value sẽ là 1 chuỗi - VD: "Giá trị của A1 nè"
; Return values .: Success	- Return sẽ là True
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_Write($range,$value)
	Local $data
	Local $url = "https://sheets.googleapis.com/v4/spreadsheets/" & $GGSheet_SpreadSheet_Id & "/values/" & $range & "?valueInputOption=RAW"
	If StringRegExp($range,"\:",0) = 1 Then
		$data = '{"values": ' & __ArrayToString($value) & '}'
	Else
		$data = '{"values": [["' & $value & '"]]}'
	EndIf
	Local $res = _HttpRequest(2,$url,_Utf8ToAnsi($data),'','',$GGSheet_Header,'PUT')
	If GGSheet_Is_Access_Token_Expired($res) = 1 Then ; Access Token đã hết hạn --> Tạo mới
		GGSheet_Create_Access_Token()
		Return GGSheet_Write($range,$value)
	EndIf

	If StringRegExp($res,'updatedRange',0) = 1 Then
		Return True
	Else
		Return SetError(1,0,$res)
	EndIf
EndFunc

Func GGSheet_ArrayToRange($aValue,$startRow = Default, $startColumn = Default) ; Đang phát triển
	If $startRow = Default Then
		$startRow = 1
	EndIf

	If $startColumn = Default Then
		$startColumn = 'A'
	Else
		$startColumn = ExcelNumberToColumn($startColumn)
	EndIf

	If $startRow = 1 Then
		Local $lastRow = UBound($aValue)
		Local $lastColumn = ExcelNumberToColumn(UBound($aValue,1))
	EndIf

	Return $startColumn & $startRow & ":" & $lastColumn & $lastRow
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_SheetNew
; Description ...: Tạo 1 sheet mới
; Syntax ........: GGSheet_SheetNew($sheet_name)
; Parameters ....: $sheet_name           - Tên sheet cần tạo
; Return values .: Success	- Return sẽ là id của sheet
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_SheetNew($sheet_name)
	Local $url = 'https://sheets.googleapis.com/v4/spreadsheets/'&$GGSheet_SpreadSheet_Id&':batchUpdate'
	Local $data = '{"requests":[{"addSheet":{"properties":{"title":"'&$sheet_name&'"}}}]}'
	Local $res = _HttpRequest(2,$url,_Utf8ToAnsi($data),'','',$GGSheet_Header,'POST')
	If GGSheet_Is_Access_Token_Expired($res) = 1 Then ; Access Token đã hết hạn --> Tạo mới
		GGSheet_Create_Access_Token()
		Return GGSheet_SheetNew($sheet_name)
	EndIf
	Local $regEx = StringRegExp($res,'sheetId": (\d+),',1)
	If @error Then
		Return SetError(1,0,$res)
	EndIf
	Return $regEx[0]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_SheetList
; Description ...: Load danh sách sheet
; Syntax ........: GGSheet_SheetList()
; Return values .: Success	- Return sẽ 1 mảng 2 chiều chứa tên và id của sheet
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_SheetList()
	Local $aResult[0][2]
	Local $url = 'https://sheets.googleapis.com/v4/spreadsheets/' & $GGSheet_SpreadSheet_Id
	Local $res = _HttpRequest(2,$url,'','','',$GGSheet_Header,'GET')
	If GGSheet_Is_Access_Token_Expired($res) = 1 Then ; Access Token đã hết hạn --> Tạo mới
		GGSheet_Create_Access_Token()
		Return GGSheet_SheetList()
	EndIf

	Local $json = _HttpRequest_ParseJSON($res)
	If $json = False Then
		Return SetError(1,0,"Không thể tạo JSON từ Response")
	EndIf

	Local $filter_sheet_id = $json.filter('$.sheets..sheetId')
	Local $sheet_ids = _HttpRequest_ParseJSON($filter_sheet_id)
	If $sheet_ids = False Then
		Return SetError(2,0,"Không tìm thấy sheet_id")
	EndIf

	Local $filter_sheet_name = $json.filter('$.sheets..title')
	Local $sheet_names = _HttpRequest_ParseJSON($filter_sheet_name)
	If $sheet_names = False Then
		Return SetError(3,0,"Không tìm thấy sheet_name")
	EndIf

	$aResult =  _Array2DCreate($sheet_names,$sheet_ids)
	If @error Then
		Return SetError(3,0,$res)
	EndIf

	Return $aResult
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_SheetFind
; Description ...: Tìm kiếm sheet_id từ tên của sheet
; Syntax ........: GGSheet_SheetFind($sheet_name)
; Parameters ....: $sheet_name          - tên sheet cần tìm kiếm
; Return values .: Success	- Return sẽ id của sheet
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_SheetFind($sheet_name)
	Local $aSheet = GGSheet_SheetList()
	If @error Then
		Return SetError(1,0,$aSheet)
	EndIf

	Local $index = _ArraySearch($aSheet,$sheet_name)
	If @error Then
		Return SetError(2,0,"Không tìm thấy Sheet_ID")
	EndIf

	Return $aSheet[$index][1]
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_SheetDelete
; Description ...: Xóa sheet theo tên
; Syntax ........: GGSheet_SheetDelete($sheet_name)
; Parameters ....: $sheet_name          - tên sheet cần xóa
; Return values .: Success	- Return sẽ id của sheet
;               .: Failure	- @error - Return sẽ là response nhận được từ server
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_SheetDelete($sheet_name)
	Local $sheet_id = GGSheet_SheetFind($sheet_name)
	If @error Then
		Return SetError(1,0,$sheet_id)
	EndIf
	Local $url = 'https://sheets.googleapis.com/v4/spreadsheets/'&$GGSheet_SpreadSheet_Id&':batchUpdate'
	Local $data = '{"requests":[{"deleteSheet":{"sheetId":'&$sheet_id&'}}]}'
	Local $res = _HttpRequest(1,$url,$data,'','',$GGSheet_Header)
	If GGSheet_Is_Access_Token_Expired($res) = 1 Then ; Access Token đã hết hạn --> Tạo mới
		GGSheet_Create_Access_Token()
		Return GGSheet_SheetDelete($sheet_name)
	EndIf
	Local $regEx = StringRegExp($res,'HTTP/1.1 200 OK',0)
	If $regEx = 0 Then
		Return SetError(2,0,$res)
	EndIf
	Return $sheet_id
EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: GGSheet_Is_Access_Token_Expired
; Description ...: Kiểm tra Access Token hết hạn chưa
; Syntax ........: GGSheet_Is_Access_Token_Expired($res)
; Parameters ....: $res                 - Response từ request
; Return values .: 1 - Đã hết hạn
;				 : 0 - Chưa hết hạn
; Author ........: Trần Hùng
; ===============================================================================================================================
Func GGSheet_Is_Access_Token_Expired($res)
	Return StringRegExp($res,'ACCESS_TOKEN_EXPIRED',0)
EndFunc

Func __ArrayToString($array)
	Return '[["' & _ArrayToString($array,'","',Default,Default,'"],["') & '"]]'
EndFunc

Func ExcelNumberToColumn($number) ; Đang phát triển
    Local $column
    While $number > 0
        $remainder = Mod($number - 1, 26)
        $column = Chr(Asc("A") + $remainder) & $column
        $number = Int(($number - $remainder) / 26)
    WEnd
    Return $column
EndFunc

Func _Make2Array($s) ; Author: https://www.autoitscript.com/forum/profile/30100-jguinch/
    Local $aLines = StringRegExp($s, "(?<=[\[,])\s*\[(.*?)\]\s*[,\]]", 3), $iCountCols = 0
    For $i = 0 To UBound($aLines) - 1
        $aLines[$i] = StringRegExp($aLines[$i], "(?:^|,)\s*(?|'([^']*)'|""([^""]*)""|(.*?))(?=\s*(?:,|$))", 3)
        If UBound($aLines[$i]) > $iCountCols Then $iCountCols = UBound($aLines[$i])
    Next
    Local $aRet[UBound($aLines)][$iCountCols]
    For $y = 0 To UBound($aLines) - 1
        For $x = 0 To UBound($aLines[$y]) - 1
            $aRet[$y][$x] = ($aLines[$y])[$x]
        Next
    Next
    Return $aRet
EndFunc