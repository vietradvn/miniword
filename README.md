# MiniWord Fork — vietradvn

Fork từ [mini-software/MiniWord](https://github.com/mini-software/MiniWord) (tag `0.9.0`), sửa 3 bugs cho `@if/@endif` và `{{TableStart:key}}/{{TableEnd:key}}`.

Sử dụng trong **ris-dotnet** (`app/print` service) — DOCX template engine cho báo cáo y tế (RIS/PACS).

## Các lỗi đã fix

| # | Bug | File | Fix |
|---|-----|------|-----|
| 1 | `@if varName` (truthy check) luôn **lật ngược** kết quả | `MiniWord.Implment.cs` | `!bool.Parse(tagValue1)` → `bool.Parse(tagValue1)` |
| 2 | `@if` bên trong `@foreach` gây **crash** — `lastEleInLoop` null | `MiniWord.Implment.cs` | Tracking lại `lastEleInLoop` sau khi `@if` remove elements + `insertBeforeEle` fallback |
| 3 | `{{TableStart:key}}{{key.field}}` — regex **bắt sai** do adjacent placeholder | `MiniWord.Implment.cs` | Strip `{{TableStart:*}}`/`{{TableEnd:*}}` trước khi regex `(?<={{).*?\..*?(?=}})` |
| 4 | Dấu chấm trong **văn bản thường** giữa hai placeholder bị hiểu thành list key | `MiniWord.Implment.cs` | `.*?` → `[^{}]*?` để match không vượt ranh giới một cặp `{{...}}` |
| 5 | Thẻ bị Word **cắt run** mà phía trước có chữ thì không được ghép lại → in ra `{{tag}}` | `MiniWord.Implment.cs` | `AvoidSplitTagText`: gom khi run chứa `{{` hoặc kết thúc bằng `{`, bỏ `TrimStart` |

### Chi tiết từng fix

**Bug 1 — `@if varName` negation:**
```csharp
// MiniWord.Implment.cs — ReplaceConditionInParagraph
// Before (sai):
if (bool.Parse(tagValue1)) removeElement = false;
else removeElement = true; // tương đương !bool.Parse
// After (đúng):
removeElement = !bool.Parse(tagValue1);
```

**Bug 2 — `@if` inside `@foreach` crash:**

`ReplaceForeachStatements` tracking `lastEleInLoop = paragraph` — khi `@if` ở cuối loop bị remove, paragraph reference bị detach khỏi document tree. `tr.InsertAfterSelf(lastEleInLoop)` crash.

Fix: tracking element mới sau mỗi `Remove()`, fallback tìm sibling cuối.

**Bug 3 — TableStart regex:**

`{{TableStart:results}}{{results.stt}}{{results.name}}{{results.result}}`

Khi `AvoidSplitTagText` merge, `InnerText` = `{{TableStart:results}}{{results.stt...}}`. Regex `(?<={{).*?\..*?(?=}})` match `TableStart:results}}{{results.stt` — bắt sai vì dấu `.` ở `results.stt`.

Fix: `Regex.Replace(innerText, @"\{\{TableStart:\w+\}\}", "")` và `\{\{TableEnd:\w+\}\}` trước khi match.

**Bug 4 — dấu chấm của văn bản thường bị hiểu thành list key:**

Cùng gốc với Bug 3 nhưng tổng quát hơn: `.*?` trong `(?<={{).*?\..*?(?=}})` **không bị chặn ở `}}`**, nên một match được phép bắt đầu sau một `{{`, chạy xuyên qua `}}`, qua văn bản thường, rồi kết thúc trước một `}}` khác. Bất kỳ dấu chấm nào của văn bản nằm giữa hai placeholder đều thành "list key" giả.

Gặp thật với mẫu siêu âm tim: một ô bảng có các tiêu đề đánh số

```
1. Van hai lá      Kiểu di động: {{kieu_di_dong}} ...
2. Van động mạch chủ   Lá van: {{la_van_dmc}} ...
3. Van động mạch phổi và ĐMP ...
4. Van ba lá ...
```

→ ba khoá giả (`...}}2`, `...}}3`, `...}}4`) → `NotSupportedException: MiniWord doesn't support more than 2 list in same row`, dù ô đó không có list nào.

| innerText | regex cũ | regex mới |
|---|---|---|
| `{{Items.Name}}` | `Items.Name` ✓ | `Items.Name` ✓ |
| `{{a}} 2. text {{b}}` | `a}} 2. text {{b` ✗ | *(không match)* ✓ |
| `{{x}}{{Items.Name}}` | `x}}{{Items.Name` ✗ | `Items.Name` ✓ |

Dòng 2 ném exception khi đủ 3 khoá giả. Dòng 3 **không** ném mà `GetObjVal` trả `null` → `continue` → cả hàng bị bỏ qua, placeholder giữ nguyên `{{...}}` trong bản in mà không có lỗi nào — nên fix này cũng chữa luôn một lỗi âm thầm.

Fix: `Regex.Matches(innerText, @"(?<=\{\{)[^{}]*?\.[^{}]*?(?=\}\})")`. Regex ở nhánh `else` (chỉ dùng `.Success`) cũng chặn tương tự.

**Bug 5 — thẻ bị cắt run mà phía trước có chữ:**

Word tách một đoạn văn thành nhiều `<w:r><w:t>` tuỳ ý (rsid, spell-check, format). `AvoidSplitTagText` sinh ra để ghép lại, nhưng chỉ khởi động khi run **mở đầu** bằng `{`:

```csharp
if (text.InnerText.TrimStart().StartsWith("{")) needAppend = true;
...
var tagContains = s.StartsWith(tagStart) && s.Contains(tagEnd);
```

Gặp thật với mẫu siêu âm tim — Word cắt thành 4 run:

```
'Họ và tên: {{ho_v'  +  'a'  +  '_ten}'  +  '}'
```

Run đầu mở đầu bằng `H` → không gom → `ReplaceText` không thấy thẻ nào nguyên vẹn → bản in giữ nguyên `{{ho_va_ten}}` dù giá trị đã có trong `vars`.

Fix:

- khởi động khi run **chứa** `{{` (thẻ bắt đầu giữa run) **hoặc kết thúc bằng** `{` (Word cắt đúng giữa hai ngoặc). Không dùng `Contains("{")` trơn — một `{` lẻ giữa câu sẽ kéo hàng loạt run không liên quan vào cùng pool rồi ép hết về định dạng của run đầu;
- `tagContains` đổi sang "có `{{` và một `}}` nằm sau nó", vì `s` giờ được phép mang phần chữ đứng trước;
- thêm lối thoát: gom quá 1000 ký tự mà chưa thành thẻ thì buông pool, tránh nuốt nốt tài liệu.

**Bỏ luôn `TrimStart()`** — nó đang xoá dấu cách đứng trước thẻ: `'Đại chỉ:'` + `' {{dia_chi}}'` in ra `Đại chỉ:123 Nguyễn Huệ`. Run ghép cũng được gắn `xml:space="preserve"` để Word không nuốt lại.

Đối chiếu 5 mẫu in thật (trước/sau): đều HTTP 200, nội dung giống hệt sau khi bỏ khoảng trắng, chỉ dài thêm 2–3 ký tự đúng bằng số dấu cách được trả lại.

## Cách build

```bash
cd src/MiniWord
dotnet pack -c Release -o /tmp/nupkg
```

## Tags hỗ trợ

| Tag | Mô tả |
|-----|-------|
| `{{key}}` | Text replacement |
| `@if @endif` | Multi-paragraph conditional |
| `@foreach @endforeach` | List iteration |
| `{{TableStart:key}} ... {{TableEnd:key}}` | Table row repeat |
| `{{#imageKey}}` | Embedded image (byte[]) |

Template phải được tạo bằng **LibreOffice / Word / OnlyOffice** — tạo programmatic thiếu XML parts cần thiết.

## Tests

```
tests/MiniWordBugFixTests/IfForeachTableTests.cs
  → 13 tests pass, 1 skip (cần .docx tạo từ Word)
```

## License

Giữ nguyên license của bản gốc (MIT).
