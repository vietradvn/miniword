# MiniWord Fork — vietradvn

Fork từ [mini-software/MiniWord](https://github.com/mini-software/MiniWord) (tag `0.9.0`), sửa 3 bugs cho `@if/@endif` và `{{TableStart:key}}/{{TableEnd:key}}`.

Sử dụng trong **ris-dotnet** (`app/print` service) — DOCX template engine cho báo cáo y tế (RIS/PACS).

## Các lỗi đã fix

| # | Bug | File | Fix |
|---|-----|------|-----|
| 1 | `@if varName` (truthy check) luôn **lật ngược** kết quả | `MiniWord.Implment.cs` | `!bool.Parse(tagValue1)` → `bool.Parse(tagValue1)` |
| 2 | `@if` bên trong `@foreach` gây **crash** — `lastEleInLoop` null | `MiniWord.Implment.cs` | Tracking lại `lastEleInLoop` sau khi `@if` remove elements + `insertBeforeEle` fallback |
| 3 | `{{TableStart:key}}{{key.field}}` — regex **bắt sai** do adjacent placeholder | `MiniWord.Implment.cs` | Strip `{{TableStart:*}}`/`{{TableEnd:*}}` trước khi regex `(?<={{).*?\..*?(?=}})` |

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
