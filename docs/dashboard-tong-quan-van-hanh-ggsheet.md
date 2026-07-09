# Thiết kế trang "Tổng quan vận hành" trên Google Sheets — PCU 2026

Thiết kế dashboard cho file `Dash_Report_PCU_2026` (Google Sheets), trang **Tổng quan vận hành**.
File mẫu dựng sẵn được sinh bởi `tools/build_tong_quan_dashboard.py` (xuất ra
`Dashboard_TongQuanVanHanh_PCU2026.xlsx` — mở thẳng bằng Google Sheets là chạy).

## 1. Yêu cầu

Trang tổng cần **cái nhìn tổng năm + so sánh tháng**, mọi báo cáo vẫn **lọc được theo thời gian**:

1. Tổng **CO Nhập** theo SPDV — theo tháng → biểu đồ **cột chồng**
2. Tổng **CO Xuất** theo SPDV — theo tháng → biểu đồ **cột chồng**
3. Tổng **CO Nhập** theo SPDV — cả kỳ → biểu đồ **tròn** (tỉ trọng)
4. Tổng **Thiệt hại** theo SPDV — theo tháng → biểu đồ **cột chồng**
5. Tổng **CO Tồn** → thẻ KPI + biểu đồ cột chồng theo tháng

## 2. Kiến trúc 3 lớp

```
'Data theo SPDV'  ──►  'DB_Tổng quan'  ──►  'Tổng quan vận hành'
 (sheet nguồn,          (lớp trung gian:      (bộ lọc + 5 thẻ KPI
  đã có sẵn)             áp bộ lọc tháng,      + 5 biểu đồ
                         đổi đơn vị triệu ₫)   + bảng so sánh tháng)
```

Lý do cần lớp trung gian: biểu đồ Google Sheets không tự lọc theo ô nhập liệu.
`DB_Tổng quan` nhận bộ lọc, **làm rỗng** các tháng ngoài kỳ lọc (`""`) → cột của
tháng đó biến mất khỏi biểu đồ; các phép SUM bỏ qua chuỗi rỗng nên tỉ trọng pie
và KPI đều tính đúng theo kỳ. Sheet này có thể **ẩn** đi sau khi dựng xong.

## 3. Vị trí dữ liệu trong sheet nguồn `Data theo SPDV`

Cột `F:Q` = tháng 1–12, cột `E` = tổng năm, nhãn SPDV ở cột `D`.

| Khối | Hàng tổng | Hàng chi tiết SPDV |
|---|---|---|
| Tổng CO Nhập | 4 | 5–13 (Robux, Token, Gift Card, Nick Random Roblox, Razer Gold, RBX Daily, RBX HQS, Topup, Supercell) |
| Tổng CO Xuất | 17 | 18–26 (cùng thứ tự, Token = "Item (Token)") |
| Thiệt hại | 35 | 36–40 (Robux die, Nick Random Roblox Die, Razer Gold, RBX (phí), Gift card lỗi không đc NCC bảo hành) |
| Tổng CO Tồn | 47 | 48–50 (Item, Gift Card (Hub), Nick Random Roblox) |

> Nếu sau này chèn/xoá hàng trong `Data theo SPDV`, cập nhật lại các hằng số hàng
> ở đầu `tools/build_tong_quan_dashboard.py` rồi chạy lại script.

## 4. Bộ lọc thời gian

- `'Tổng quan vận hành'!C4` = **Từ tháng**, `E4` = **Đến tháng** (số 1–12, có Data validation).
- Mặc định `1 → 12` = cái nhìn tổng cả năm; thu hẹp để soi giai đoạn.
- `DB_Tổng quan!B1`/`C1` soi gương 2 ô này: `='Tổng quan vận hành'!C4`.

## 5. Sheet `DB_Tổng quan` (đơn vị: triệu ₫)

5 khối, mỗi khối tháng gồm: hàng tiêu đề, 12 hàng `Th1..Th12`, cột `Tổng`, hàng `Tổng kỳ lọc`.

| Khối | Vùng | Nguồn |
|---|---|---|
| 1. CO Nhập theo SPDV | header hàng 4, dữ liệu `A5:K16`, tổng hàng 17 | `F5:Q13` |
| 2. CO Xuất theo SPDV | header hàng 20, dữ liệu `A21:K32`, tổng hàng 33 | `F18:Q26` |
| 3. Tỉ trọng CO Nhập (pie) | `A36:B45` — mỗi SPDV = tổng cột của khối 1 (`=B17`…`=J17`) | khối 1 |
| 4. Thiệt hại theo nhóm | header hàng 48, dữ liệu `A49:G60`, tổng hàng 61 | `F36:Q40` |
| 5. CO Tồn theo nhóm | header hàng 64, dữ liệu `A65:E76`, tổng hàng 77 | `F48:Q50` |

Công thức mẫu 1 ô (tháng `m`, ví dụ Robux tháng 1 tại `B5`):

```
=IF(AND(1>=$B$1,1<=$C$1),'Data theo SPDV'!F5/1000000,"")
```

Trên Google Sheets có thể thay cả khối 12×9 bằng **một** công thức mảng (đặt tại `B5`):

```
=ARRAYFORMULA(IF((SEQUENCE(12)>=$B$1)*(SEQUENCE(12)<=$C$1),
  TRANSPOSE('Data theo SPDV'!F5:Q13)/1000000, ))
```

## 6. Trang `Tổng quan vận hành`

### 6.1 Thẻ KPI (hàng 6–8, đơn vị ₫)

Tính thẳng từ hàng tổng của sheet nguồn theo bộ lọc — mẫu (CO Nhập, hàng 4):

```
=SUMPRODUCT('Data theo SPDV'!$F$4:$Q$4,
  (COLUMN('Data theo SPDV'!$F$4:$Q$4)-5>=$C$4)*(COLUMN('Data theo SPDV'!$F$4:$Q$4)-5<=$E$4))
```

| Thẻ | Hàng tổng nguồn |
|---|---|
| Tổng CO Nhập | 4 |
| Tổng CO Xuất | 17 |
| Chênh lệch Nhập − Xuất | `=B7-E7` |
| Tổng Thiệt hại (+ ghi chú `% trên CO Nhập`) | 35 |
| Tổng CO Tồn (cộng dồn kỳ — cùng logic cột tổng năm của sheet nguồn) | 47 |

### 6.2 Năm biểu đồ (Insert ▸ Chart trên Google Sheets)

| # | Biểu đồ | Loại | Vùng dữ liệu (`DB_Tổng quan`) | Nhãn cột (X) |
|---|---|---|---|---|
| 1 | CO Nhập theo SPDV theo tháng | Cột **chồng** (Stacked column) | `A4:J16` | `A5:A16` |
| 2 | CO Xuất theo SPDV theo tháng | Cột chồng | `A20:J32` | `A21:A32` |
| 3 | Tỉ trọng CO Nhập theo SPDV kỳ lọc | **Tròn** (Pie), nhãn = % | `A36:B45` | — |
| 4 | Thiệt hại theo SPDV theo tháng | Cột chồng | `A48:F60` | `A49:A60` |
| 5 | CO Tồn theo SPDV theo tháng | Cột chồng | `A64:D76` | `A65:A76` |

Cấu hình chung: "Use row 4/20/… as headers", legend dưới biểu đồ, trục tung định dạng `#,##0`.

### 6.3 Bảng "So sánh tháng" (L43:Q57)

Tháng | CO Nhập | **% MoM Nhập** | CO Xuất | Thiệt hại | CO Tồn — lấy từ cột `Tổng` các
khối DB, chỉ hiện tháng trong kỳ lọc; hàng cuối là tổng kỳ. Đây cũng là "table view"
cho người khó phân biệt màu.

## 7. Bảng màu SPDV (cố định theo thực thể trên mọi biểu đồ)

Palette categorical đã kiểm tra an toàn mù màu (worst adjacent CVD ΔE 24.2 — PASS):

| SPDV | Hex | | Nhóm thiệt hại / tồn | Hex |
|---|---|---|---|---|
| Robux | `#2A78D6` | | Robux die | `#2A78D6` |
| Token / Item | `#898781` | | Nick Random Roblox Die | `#EDA100` |
| Gift Card (mọi biến thể) | `#1BAF7A` | | Razer Gold | `#008300` |
| Nick Random Roblox | `#EDA100` | | RBX (phí) | `#4A3AA7` |
| Razer Gold | `#008300` | | Gift card lỗi | `#1BAF7A` |
| RBX Daily | `#4A3AA7` | | Item (tồn) | `#898781` |
| RBX HQS | `#E34948` | | | |
| Topup | `#E87BA4` | | | |
| Supercell | `#EB6834` | | | |

Nguyên tắc: màu đi theo **thực thể** (Gift Card luôn xanh ngọc ở cả 5 biểu đồ),
không đi theo vị trí series; nhóm bằng 0 cả năm (Token) nhận màu xám trung tính.

## 8. Đưa vào file Google Sheets thật

**Cách A — khuyên dùng (copy 2 sheet):**
1. Tải `Dashboard_TongQuanVanHanh_PCU2026.xlsx` lên Google Drive, mở bằng Google Sheets.
2. Chuột phải tab **`Tổng quan vận hành`** ▸ *Sao chép sang* ▸ *Bảng tính hiện có* ▸ chọn file PCU thật.
3. Làm tương tự với tab **`DB_Tổng quan`**.
4. Mở file thật, đổi tên 2 sheet nếu bị thêm hậu tố "Bản sao của …". Công thức tự bám vào
   `Data theo SPDV` thật (trùng tên sheet). Nếu biểu đồ nào mất vùng dữ liệu, sửa lại theo bảng §6.2 (≈1 phút/biểu đồ).

**Cách B — dashboard tách file, tự cập nhật:** giữ nguyên file dựng sẵn, thay sheet snapshot
`Data theo SPDV` bằng một sheet cùng tên chỉ có 1 công thức tại `A1`:

```
=IMPORTRANGE("<URL file PCU thật>","Data theo SPDV!A1:R50")
```

(bấm *Allow access* lần đầu) → dashboard luôn ăn theo số liệu thật mà không đụng file gốc.

## 9. Hoàn thiện sau khi lắp

- Ẩn sheet `DB_Tổng quan` (chuột phải tab ▸ Ẩn trang tính).
- `Data > Protect range` cho toàn trang trừ `C4`, `E4` để người xem chỉ đổi được bộ lọc.
- Có thể thay ô `C4/E4` bằng Dropdown chips (Insert ▸ Dropdown) danh sách 1–12 cho đẹp.
