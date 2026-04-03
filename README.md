# VBA DataFrame

PythonのPandas DataFrameに似た操作をVBAで実現するクラスモジュールです。  
メソッドチェーンによる直感的な操作が可能です。

## セットアップ

### インポート手順

1. Excel VBAエディタを開く（`Alt+F11`）
2. メニュー → **ファイル** → **ファイルのインポート**
3. 以下の3ファイルを順にインポート：
   - `src/DataFrame.cls`
   - `src/DataFrameGroupBy.cls`
   - `src/DFrame.bas`
   - clsのインポートがうまくいかない場合は、改行をLFからCRLFに変更してインポートし直してください。
4. （任意）`examples/Example_Basic.bas` もインポート

> **Note:** 外部参照の追加は不要です（`Scripting.Dictionary` は late binding で使用）。

## クイックスタート

```vba
' データ作成
Dim df As DataFrame
Set df = DFrame.Create( _
    Array("Name", "Age", "City", "Sales"), _
    Array("Tanaka", 30, "Tokyo", 15000), _
    Array("Suzuki", 25, "Osaka", 8000), _
    Array("Sato", 35, "Tokyo", 22000))

' 表示
df.Print

' メソッドチェーン
df.Where("Sales", ">", 10000) _
  .OrderBy("Sales", Ascending:=False) _
  .Sel("Name", "Sales") _
  .Print
```

## ファクトリ関数 (`DFrame` モジュール)

| 関数 | 説明 |
|---|---|
| `DFrame.FromRange(rng, [hasHeader])` | Excel Range から作成 |
| `DFrame.FromArray(data2D, colNames)` | 2D配列から作成 |
| `DFrame.Create(colNames, row1, row2, ...)` | 行データを直接指定 |
| `DFrame.FromCsv(filePath, [sep], [hasHeader])` | CSVファイルから作成 |
| `DFrame.EmptyFrame(colNames)` | 空の DataFrame を作成 |

## プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| `.RowCount` | `Long` | 行数 |
| `.ColCount` | `Long` | 列数 |
| `.Shape` | `String` | `"N rows x M cols"` |
| `.Columns` | `Variant` | カラム名配列 |
| `.Value(row, col)` | `Variant` | セル値（col は名前 or インデックス）|
| `.Values` | `Variant` | 内部2D配列 |

## データアクセス

| メソッド | 戻り値 | 説明 |
|---|---|---|
| `.Col(colName)` | `Variant` | 列を1D配列で取得 |
| `.Row(rowIdx)` | `Variant` | 行を1D配列で取得 |

## 選択・フィルタ

| メソッド | 説明 |
|---|---|
| `.Head(n)` | 先頭 n 行 |
| `.Tail(n)` | 末尾 n 行 |
| `.Sel(col1, col2, ...)` | 列の選択 |
| `.Where(col, op, value)` | 条件フィルタ |
| `.Slice(start, end)` | 行範囲のスライス |
| `.Distinct([colName])` | 重複削除 |

### Where 演算子

| 演算子 | 説明 | 例 |
|---|---|---|
| `=` | 等しい | `.Where("City", "=", "Tokyo")` |
| `<>` | 等しくない | `.Where("Age", "<>", 30)` |
| `>`, `>=`, `<`, `<=` | 大小比較 | `.Where("Sales", ">", 10000)` |
| `Like` | パターンマッチ | `.Where("Name", "Like", "T*")` |
| `In` | 配列内に含まれる | `.Where("City", "In", Array("Tokyo","Osaka"))` |

## 変換

| メソッド | 説明 |
|---|---|
| `.OrderBy(col, [ascending])` | ソート |
| `.AddCol(name, values)` | 列追加（配列 or スカラー）|
| `.RemoveCol(colName)` | 列削除 |
| `.RenameCol(old, new)` | 列名変更 |

## 集計

| メソッド | 説明 |
|---|---|
| `.Sum(col)` | 合計 |
| `.Mean(col)` | 平均 |
| `.MaxVal(col)` | 最大値 |
| `.MinVal(col)` | 最小値 |
| `.Describe()` | 基本統計量（Count/Sum/Mean/Min/Max）|

## GroupBy

```vba
' 地域別の売上合計
df.GroupBy("Region").Sum("Sales").Print

' 複数キーでグループ化
df.GroupBy("Region", "Product").Mean("Sales").Print

' 件数カウント
df.GroupBy("Region").Count.Print
```

| メソッド | 説明 |
|---|---|
| `.Sum(col1, col2, ...)` | グループ別合計 |
| `.Mean(col1, col2, ...)` | グループ別平均 |
| `.Count` | グループ別件数 |
| `.MaxVal(col1, ...)` | グループ別最大値 |
| `.MinVal(col1, ...)` | グループ別最小値 |

## 結合

```vba
' Inner Join
orders.JoinDF(customers, "CustID", "inner").Print

' Left Join
orders.JoinDF(customers, "CustID", "left").Print

' 縦結合
df1.VStack(df2).Print
```

## 出力

| メソッド | 説明 |
|---|---|
| `.ToRange(targetCell)` | Range に書き出し（ヘッダー付き）|
| `.ToArray()` | 2D配列として返す |
| `.Print([maxRows])` | Immediate Window に表示 |
| `.ToString([maxRows])` | 整形済み文字列を返す |

## テスト

`tests/TestDataFrame.bas` をインポート後、Immediate Window で `TestAll` を実行：

```
TestAll
```

## ファイル構成

```
vba_dataframe/
├── src/
│   ├── DataFrame.cls           ← メインクラス
│   ├── DataFrameGroupBy.cls    ← GroupBy ヘルパー
│   └── DFrame.bas              ← ファクトリ関数
├── examples/
│   └── Example_Basic.bas       ← 使用例
├── tests/
│   └── TestDataFrame.bas       ← テストスイート
└── README.md
```

## 動作要件

- Excel 2010 以降（推奨: Excel 2016+）
- 外部参照の追加は不要
