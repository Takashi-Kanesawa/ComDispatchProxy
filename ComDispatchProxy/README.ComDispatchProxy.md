# ComDispatchProxy (NuGet: `ComDispatchProxy`)

`DispatchProxy` を使って **COM オブジェクトをプロキシ化**し、呼び出しの中継と **COM の解放（`Marshal.ReleaseComObject`）**をまとめて扱うための小さなライブラリです。

> ⚠️ 現時点では **Microsoft Office Interop Excel でのみ動作確認**しています。  
> Excel 以外の COM（Word/PowerPoint/独自COM等）でも動く設計にはしてありますが、**未検証**です。

---

## 動作環境

- .NET: **net8.0**
- 実行環境: **Windows 推奨（COM 前提）**
  - 本ライブラリは `Marshal.ReleaseComObject` / `QueryInterface` を使用します（非Windowsでは動作しない可能性が高いです）

---

## 何を解決するか

COM Interop（特に Office）でありがちな

- `System.__ComObject` しか見えず、型判定が難しい
- どこで RCW が残っているか分からず、解放漏れしがち
- 親子の COM オブジェクト（例: Application → Workbooks → Workbook …）を手動で全部 `ReleaseComObject` するのが面倒

といった問題に対して、

- COM のメソッド呼び出しを `DispatchProxy` で中継
- 返り値が COM オブジェクトなら **自動でプロキシ化**
- 親プロキシが `Dispose` されると、子プロキシもまとめて `Dispose` → `ReleaseComObject`

という流れを提供します。

---

## 使い方の全体像

このライブラリ単体では「どの COM インターフェイス型をプロキシ対象にするか」を知りません。  
そのため、**`IComProxyFactory`（通常は `ComProxyFactoryBase` 派生）**を用意して使います。

- `ComDispatchProxy<T>`  
  COM 呼び出しを中継し、返ってきた COM を `IComProxyFactory` 経由で再ラップします
- `ComProxyFactoryBase`  
  `usedTypes`（プロキシ化したいインターフェイスの集合）を受け取り、`QueryInterface` で一致判定して `ComDispatchProxy<ThatInterface>` を生成します

---

## 最小サンプル（独自 COM 向けの雛形）

```csharp
using ComDispatchProxy;

public sealed class MyComProxyFactory : ComProxyFactoryBase
{
    protected override string TargetAssemblyName => "Your.Interop.AssemblyName";

    public MyComProxyFactory(IEnumerable<Type> usedTypes) : base(usedTypes) { }

    // まずは「インターフェイスだけ」を対象にするのが無難
    protected override IEnumerable<Type> FilteredTypes =>
        base.FilteredTypes.Where(t => t.IsInterface);
}
// --- 利用側（例） ---
// 基本: 対象 COM の Interop アセンブリの GetTypes() を渡す
// var usedTypes = typeof(YourRootInterface).Assembly.GetTypes();
// var factory = new MyComProxyFactory(usedTypes);
// using var root = ComDispatchProxy<YourRootInterface>.CreateProxy(factory, CreateComInstanceSomehow());
// var api = root.Proxy;
```

> ✅ usedTypes は 基本的に「対象 COM の Interop アセンブリの Assembly.GetTypes()」をそのまま渡すことを想定しています。  
> その前提では、通常、ライブラリ経由で取得する COM オブジェクトは usedTypes のいずれかに一致し、プロキシ化されます。  
> ⚠️ ただし、何らかの理由で usedTypes に一致する型が見つからない場合、ComProxyFactoryBase.CreateProxyByFactoryFunction() は 元の COM オブジェクトをそのまま返します。  
> このケースは「例外パス」であり、そのオブジェクトは 自動解放ツリーの管理対象外になります（必要なら呼び出し側で解放してください）。  
> ※ usedTypes に余分な型が含まれている場合は、FilteredTypes 側で対象を限定してください。

> ✅ `ComProxyFactoryBase.CreateProxyByFactoryFunction()` は、`usedTypes` に一致する型が見つからなければ **元の COM オブジェクトをそのまま返します**。  
> つまり、**`usedTypes` に含めていない型で返ってきた COM は自動解放の管理対象外**になります（必要なら自分で解放してください）。

---

## ログ

デバッグ用ログを ON/OFF できます。

```csharp
using ComDispatchProxy;

ComProxyLogConfig.Enabled = true;
ComProxyLogConfig.Writer  = msg => Console.WriteLine(msg); // 任意（未指定なら DEBUG 時に Debug.WriteLine）
```

---

## 設計メモ（挙動）

- メソッド呼び出し時:
  - 引数に `IComDispatchProxy` が含まれていたら、内部の **生 COM（`RowObject`）**に戻して呼び出します
  - 返り値が COM の場合:
    - `ProxyFactory.CreateProxyByFactoryFunction(result, this)` で子プロキシ化
    - 子が `IComDispatchProxy` なら親子関係を登録
- `Dispose`:
  - 子プロキシを先に `Dispose`
  - 自身の COM を `Marshal.ReleaseComObject` で解放
- ルートプロキシ（親が `null` のもの）については、プロセス終了/未処理例外のタイミングで `DisposeIfRoot()` を呼ぶガードも入っています  
  - ただし **基本は `using` で明示的に破棄**してください（ガードは保険です）

---

## 既知の注意点 / 制限

- **Excel 以外は未検証**です（設計上は汎用ですが、動作実績は Excel のみ）
- `usedTypes` に含まれない COM は自動管理されません
- COM の参照が別経路で保持されている場合（生 COM をどこかに保持した等）は、当然ながら解放できません  
  → 「プロキシを通して触る」運用が前提です

