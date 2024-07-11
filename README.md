# LinkedWindowToRange
 Class module bổ trợ hiển thị Userform tại ô Excel
 
Với lớp LinkedWindowToRange chúng ta có thể dễ dàng đặt cửa sổ vào đúng vị trí tại ô đang chọn một cách dễ dàng

### Ưu điểm của lớp:
- Đặt cửa sổ vào vị trí chính xác, kể cả chế độ hiển thị phải sang trái.
- Dễ dàng tạo mã VBA để sử dụng lớp.
- Tương thích khi sử dụng nhiều màn hình.
- Tự động cân chỉnh vị trí khi cửa sổ nằm ngoài màn hình hiển thị.
- Nhúng cửa sổ vào cửa sổ Excel7, làm cho Excel7 thành cửa sổ phụ thuộc.
- Đặt tiêu đề tiếng Việt cho Form, ẩn hoặc hiện tiêu đề.
- Làm mờ cửa sổ.

### Hướng dẫn sử dụng lớp LinkedWindowToRange:

#### Khởi tạo lớp
Ta có thể khai báo toàn cục để truy cập lại, hoặc cục bộ để sử dụng 1 lần duy nhất

#### Phương thức trong lớp

| Phương thức      | Phương thức                          | Kiểu    | Diễn giải                                                                                     |
| ---------------- | ------------------------------------ | ------- | --------------------------------------------------------------------------------------------- |
| **newForm**      | set LinkWindow.newForm = userform1   | Object  | Nhập userform để lớp khởi tạo cho form này                                                    |
| **newWindow**    | LinkWindow.newWindow = hwnd          | Long    | Nhập hwnd cửa sổ để lớp khởi tạo cho cửa sổ này<br>(Nếu không sử dụng form, có thể nhập hwnd) |
| **LinkedWindow** | LinkWindow.LinkedWindow = True       | Boolean | Nhúng cửa sổ vào cửa sổ Excel7, buộc cửa sổ phụ thuộc cửa sổ Excel 7                          |
| **Offset**       | LinkWindow.Offset 5, 5               | Long    | Xê dịch cửa sổ                                                                                |
| **Show**         | LinkWindow.Show Range("A1"), 4       |         | Hiển thị cửa sổ tại ô, và kiểu hiển thị                                                       |
| **ReShow**       | LinkWindow.ReShow                    |         | Hiển thị lại cửa sổ, sau khi đã đổi các thiết đặt                                             |
| **SetTitle**     | LinkWindow.SetTitle "Tiêu đề cửa sổ" |         | Đặt tiêu đề cho cửa sổ (Hỗ trợ ký tự việt)                                                    |
| **ShowTitle**    | LinkWindow.ShowTitle                 |         | Hiển thị tiêu đề cửa sổ                                                                       |
| **HideTitle**    | LinkWindow.HideTitle                 |         | Ẩn thị tiêu đề cửa sổ                                                                         |
| **Transparent**  | LinkWindow.Transparent 0.7           |         | Tạo độ mờ cho cửa sổ  


#### Kiểu hiển thị cho phương thức Show

Vị trí bắt đầu​​ | Diễn giải​
--------------|-----------------
RPE_leftTop = 0​​ | Vị trí bên trái + phía trên​
RPE_leftBottom = 1​​ | Vị trí bên trái + phía dưới​
RPE_RightTop = 2​​ | Vị trí bên phải + phía trên​
RPE_RightBottom = 4​​ | Vị trí bên phải + phía dưới​

Vị trí cửa sổ ​| Diễn giải​
--------------|-----------------
RPE_WindowRightBelow = 0​ | Cửa sổ sẽ nằm ở bên phải + phía dưới​
RPE_WindowRightAbove = 2 ^ 3​​ | Cửa sổ sẽ nằm ở bên phải + phía trên​
RPE_WindowLeftBelow = 2 ^ 4​​ | Cửa sổ sẽ nằm ở bên trái + phía dưới​
RPE_WindowLeftAbove = 2 ^ 5​​ | Cửa sổ sẽ nằm ở bên trên + phía trên​
RPE_FullScreen = 2 ^ 6​​ | Hiển thị toàn màn hình​

Mặc định cửa sổ sẽ hiển thị RPE_leftTop + RPE_WindowRightBelow
Ví dụ: ​
```vba
LinkWindow.Show Range("A1"), 0
```

Thử đặt cửa sổ vị trí bên phải + phía trên ô và cửa sổ nằm bên trái + phía trên
Ví dụ:
```vba
LinkWindow.Show Range("A1"), RPE_RightTop + RPE_WindowLeftAbove
```
Ví dụ mã đầy đủ cho các phương thức

```vba
Sub ShowForm()
  Dim LinkWindow As LinkedWindowToRange
  Set LinkWindow = New LinkedWindowToRange
  With LinkWindow
     Set .newForm = formRangePosition
    .SetTitle "Tiêu đề"
    .hideTitle
    .linkedWindow = True
    .Offset 0, 0
    .Transparent 0.7
    .Show [B3], RPE_LeftTop + RPE_WindowRightBelow
  End With
End Sub```
