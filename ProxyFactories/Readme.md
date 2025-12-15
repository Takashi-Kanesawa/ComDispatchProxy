# ComDispatchProxy.Excel (NuGet: `ComDispatchProxy.Excel`)

`ComDispatchProxy` の **Excel Interop 向け**の補助パッケージです。  
`InteropExcelProxyFactory` を提供します（内部的には `ComProxyFactoryBase` を継承）。

---

## できること

- `Microsoft.Office.Interop.Excel` で返ってくる COM オブジェクトを **自動でプロキシ化**
- `using` でルート（`Excel.Application`）を破棄すれば、子プロキシもまとめて `ReleaseComObject`

---

## 動作環境

- .NET: **net8.0**
- OS: **Windows**
- 実行環境に **Microsoft Excel（デスクトップ版）がインストールされていること**
  - `Microsoft.Office.Interop.Excel` を介して COM オートメーション（`new Excel.Application()`）を行います
  - Excel がインストールされていない環境では動作しません
---

## 使い方（最小サンプル）

```csharp
using Excel = Microsoft.Office.Interop.Excel;
using ComDispatchProxy;
using ComDispatchProxy.ProxyFactories;

ComProxyLogConfig.Enabled = true;

var usedTypes = typeof(Excel.Application).Assembly.GetTypes();
var excelFactory = new InteropExcelProxyFactory(usedTypes);

using (var excelAppRoot =
       ComDispatchProxy<Excel.Application>.CreateProxy(excelFactory, new Excel.Application()))
{
    var excelApp = excelAppRoot.Proxy;

    var wb = excelApp.Workbooks.Add();
    var ws = (Excel.Worksheet)wb.Worksheets[1];

    ws.Cells[1, 1].Value = "Hello";
}
```

---

## 重要: `usedTypes` の考え方（設計意図と推奨運用）

この Factory は、COM の返り値をプロキシ化する際に `QueryInterface` を用いた型判定を行います。
その判定コストを抑えるため、`usedTypes` には **「実際にプロキシ化したい Interop 型だけ」** を渡す設計になっています（Excel の全インターフェイス一覧を目指しません）。

```csharp
var usedTypes = typeof(Excel.Application).Assembly.GetTypes();
var factory   = new InteropExcelProxyFactory(usedTypes);
```

### 推奨: Excel Interop を扱うアセンブリは 1 つに集約する

`usedTypes` は **`GetTypes()` を呼んだアセンブリ内の型集合**になります。
そのため、アプリケーションが複数 DLL 構成の場合でも、**Excel Interop を参照・利用するコードは 1 つのアセンブリ（1 DLL / 1 EXE）に集約し、そのアセンブリで `GetTypes()` する**運用を推奨します。

> 複数アセンブリにまたがって Excel Interop を利用し、各アセンブリの `usedTypes` を統合する運用は **現時点では未検証**です。

### `usedTypes` に含まれない型が返ってきた場合

返ってきた COM オブジェクトが `usedTypes` のいずれにも一致しない場合、Factory は **その COM をプロキシ化せず、そのまま返します**。
（この場合、その COM は本ライブラリの自動解放ツリーに参加しません。必要に応じて呼び出し側で解放してください）


---

## `_` 付きの型を除外している理由

`InteropExcelProxyFactory` では、`usedTypes` から **`_` で始まるインターフェイス型**を除外しています。

```csharp
protected override IEnumerable<Type> FilteredTypes =>
    base.FilteredTypes.Where(t =>
        t.IsInterface &&
        (t.Name.StartsWith("_", StringComparison.Ordinal)) == false);
```

Interop.Excel の環境によっては `_Application` など **`_` 付きの型が混ざる**ことがあり、
これらは実用上の型判定・プロキシ生成の対象として扱わない方が安定するため、Factory 内でフィルタしています。

---

## Tips: `object` が返る API と型の勘違いに注意

Excel Interop には、コレクションの `Item`（インデクサ）などが `object` を返す箇所があります（例: `Workbook.Worksheets` は型として `Excel.Sheets` を返し、`Sheets[1]` は `object` です）。

`object` のままだと IntelliSense が効きづらいだけでなく、**「返ってくる実体の型を勘違いしている」ことに気づきにくい**のが地味に危険です。
（例えば `Sheets` には `Worksheet` だけでなく `Chart` が含まれる場合があり、`(Excel.Worksheet)sheets[1]` が実行時に失敗することがあります）

そのため、必要に応じて明示キャスト（または型チェック）してください。

```csharp
var sheets = wb.Worksheets;                  // 型は Excel.Sheets
var ws = (Excel.Worksheet)sheets[1];         // object → Worksheet に明示キャスト
// あるいは安全寄りに:
// if (sheets[1] is Excel.Worksheet ws) { ... }
```

---

## 未検証事項

- この Factory を Excel 以外（Word/PowerPoint 等）に転用した場合の動作
