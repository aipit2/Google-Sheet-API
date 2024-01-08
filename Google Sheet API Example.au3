#include <include/_HttpRequest.au3>
#include <Google Sheet API.au3>

Local $spread_sheet_id = "1ALGrcV38M-4Sdmm0BBsQ7QLTncnoom8nxOKh6-3Hdf0" ; Thay thể bằng spreadsheetid của bác
Local $service_account_file_dir = @ScriptDir & "\key.json" ; Thay thế thành đường dẫn tới file file serivce account của các bác
Local $library_file_dir = @ScriptDir & '\include\jsrsasign.js' ; Thay thế thành đường dẫn tới file thư viện của các bác

; Hàm này dùng để khai báo
Local $access_token = GGSheet_Setup($service_account_file_dir,$library_file_dir, $spread_sheet_id)
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$access_token)
	Exit
EndIf

; Ví dụ đọc dữ liệu 1 ô --> Giá trị trả về sẽ là 1 chuỗi
Local $value = GGSheet_Read("Sheet1!A1")
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$value)
	Exit
EndIf
MsgBox(64 + 8192 + 262144,"Đọc giá trị từ 1 ô",$value)

; Ví dụ đọc dữ liệu nhiều ô --> Giá trị trả về sẽ là 1 mảng 2 chiều
Local $value = GGSheet_Read("Sheet1!A1:B4")
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$value)
	Exit
EndIf
_ArrayDisplay($value,"Đọc dữ liệu từ vùng: Sheet1!A1:B4","",0,Default,"A|B")

; Ví dụ ghi dữ liệu vào 1 ô
Local $value = "Tôi muốn ghi giá trị này vào ô A1" ; Nếu muốn ghi dữ liệu vào 1 ô thì giá trị là 1 chuỗi
Local $write = GGSheet_Write("Sheet1!A1",$value)
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$write)
	Exit
EndIf

MsgBox(64 + 8192 + 262144,"Thông báo","Ghi dữ liệu thành công")

;~ ; Ví dụ ghi dữ liệu vào nhiều ô
Local $value = [["A1 nè!","B1 nè!","C1 nè!"]] ; Nếu muốn ghi dữ liệu vào nhiều ô thì giá trị là 1 mảng 2 chiều
Local $write = GGSheet_Write("Sheet1!A1:C1",$value)
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$write)
	Exit
EndIf

MsgBox(64 + 8192 + 262144,"Thông báo","Ghi dữ liệu thành công")

; Đoạn này sẽ có 1 lưu ý khi các bác ghi dữ liệu vào nhiều ô
; Như ví dụ ở trên thì mình muốn ghi dữ liệu vào vùng A1:C1 tương đương với A1,B1,C1
; Trên ví dụ mình nhập vùng muốn ghi là Sheet1!A1:C1, nhưng mình cũng có thể ghi là Sheet1!A1:C1, Sheet1!A1:C2, Sheet1!A1:C1000
; Hoặc có thể là Sheet1!A1:Z1000, miễn sao cái vùng mà các bác muốn ghi nó phải >= cái mảng chứa giá trị các bác muốn ghi
; Cái này các bác sẽ gặp khi các bác muốn ghi nhưng chỉ biết ô đầu tiên
; Thì các bác có thể để thứ 2 là Z1000 hay cột nào tùy thích, miễn là nó phải >= cái mảng chứa giá trị các bác muốn ghi

; Ví dụ ghi dữ liệu vào nhiều ô nhưng range bị nhỏ hơn số phần tử của mảng --> gây lỗi
Local $value = [["A1 nè!","B1 nè!","C1 nè!"]] ; Nếu muốn ghi dữ liệu vào nhiều ô thì giá trị là 1 mảng 2 chiều
Local $write = GGSheet_Write("Sheet1!A1:B1",$value) ; Ở đây đáng lẽ mình phải ghi là C >= 1 và cột nào >= C
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$write)
EndIf

; Ví dụ lấy danh sách tên sheet + sheet id
Local $aSheet = GGSheet_SheetList()
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$aSheet)
	Exit
EndIf

_ArrayDisplay($aSheet,"Danh sách sheet","",0,Default,"Tên sheet|Sheet ID")

; Ví dụ tạo 1 sheet mới
Local $sheet_id = GGSheet_SheetNew("Tạo sheet mới nè 2")
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$sheet_id)
	Exit
EndIf

MsgBox(64 + 8192 + 262144,"Thông báo","Tạo sheet mới thành công - Sheet ID: " & $sheet_id)

; Ví dụ xóa sheet
Local $sheet_id = GGSheet_SheetDelete("Tạo sheet mới nè 2")
If @error Then
	MsgBox(16 + 8192 + 262144,"Thông báo lỗi",$sheet_id)
	Exit
EndIf

MsgBox(64 + 8192 + 262144,"Thông báo","Xóa sheet thành công - Sheet ID: " & $sheet_id)
