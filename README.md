# XArray

 * VBAで簡単に複数のデータを扱うための配列クラスです。
 * Windows、MacいずれのVBAでも動作します。

## 特徴

 * どんな型のデータでも格納できます。複数の型を混在させることもできます。
 * ソート機能があります。
 * ソートはカスタムクラスを用意することで任意の順に並べ替えることができます。
 * APIはVBAのCollection、.Net FrameworkのArray、Objective-C (CoreFoundation)のNSArrayなどを参考に、VBAらしくて使いやすくなるよう心がけました。
 * Windows、MacいずれのVBAでも動作します。
 * クラスとして実装し、グローバル変数を使っていません。他のモジュールに干渉することを気にせず導入できます。

## インストール

Visual Basic EditorよりXArray.clsをインポートしてください。

## 使い方

 1. XArrayをインスタンス化します。
 2. Addメソッドでデータを追加します。
 3. Removeメソッドでデータを削除します。
 4. Countプロパティでデータの個数を確認します。
 5. Itemプロパティで格納したデータを取り出します。
 6. Sortメソッドでデータを並び替えます。
 7. 任意の順でソートするにはCompareメソッドを備えたカスタムクラスを作り、インスタンスをSortメソッドの引数に指定します。

## API

	Add(Value) メソッド
	Value: 追加するデータ
データを追加します。Valueはどんな型でも（オブジェクト型でも）指定できます。
********************************
	Count As Long プロパティ（読み取り専用）
格納したデータの個数を返します。何もデータがないときは0になります。
********************************
	Item(Index As Long) プロパティ（読み取り・書き込み）
	Index: 取り出す/置き換えるデータの位置
格納したデータを取り出したり、データを置き換えます。
********************************
	CompareMode As VbCompareMethod プロパティ
ExistsプロパティやSortメソッドなどでComparerを省略した場合の比較方法を取得・設定します。
********************************
	IndexOf(Value, Optional Comparer) As Long プロパティ（読み取り専用）
	Value: 探すデータ
	Comparer: データ判定を行うオブジェクト。省略可。
	戻り値: 見つかった位置。見つからなかった場合には-1を返します。
Valueで指定したデータを探し、見つかった位置を返します。Comparerを省略時はCompareModeプロパティの設定に従ってデータの判定を行います。
********************************
	Exists(Value, Optional Comparer) As Boolean プロパティ（読み取り専用）
	Value: 探すデータ
	Comparer: データ判定を行うオブジェクト。省略可。
	戻り値: データが存在すればTrue、存在しなければFalseになります。
Valueで指定したデータが存在すればTrue、存在しなければFalseになります。Comparerを省略時はCompareModeプロパティの設定に従ってデータの判定を行います。実装はIndexOfを利用しています。
********************************
	Insert(Index As Long, Value As Variant) メソッド
	Index: 挿入位置。0≦Index≦Countの範囲である必要があります。
	Value: 挿入するデータ
データを指定位置に挿入します。挿入された位置およびそれより後ろにあるデータすべての位置が1つずつ後ろにずれます。
********************************
	Remove(Index As Long) メソッド
	Index: 削除するデータの位置
指定した位置にあるデータを削除します。削除された位置より後ろにあるデータすべての位置が1つずつ前にずれます。
********************************
	Exchange(Index1 As Long, Index2 As Long) メソッド
	Index1: 入れ替えるデータの位置1
	Index2: 入れ替えるデータの位置2
Index1とIndex2の位置にあるデータを入れ替えます。
********************************
	Reverse メソッド
データの並び順を逆転します。
********************************
	Clone As XArray メソッド
	戻り値: 複製したXArrayインスタンス
XArrayインスタンスを複製し、新たなXArrayインスタンスを返します。
********************************
	Sort(Optional Comparer) メソッド
	Comparer: データ判定を行うオブジェクト。省略可。
データ位置の並べ替え（ソート）を行います。Comparerを省略時はCompareModeプロパティの設定に従ってデータの判定を行います。
********************************
	Items As Variant プロパティ（読み取り専用）
	戻り値: VBAの配列(Array)
XArrayで管理しているデータをVBAの配列(Array)として取り出します。For Eachを使いたいときやExcelにデータを書き込みたいときなどに使えます。

## Comparerについて
 * いくつかのメソッドにおいてComparerという引数があります。ここに次の要件を満たすオブジェクトを指定することで、ソートやデータが同じであることの判定などを自分で設定することができます。
   * Compareメソッドを備えたクラスのインスタンスであること。
   * Compareメソッドは引数を2つとること。
   * Compareメソッドは2つの引数を比較し、1つ目が2つ目よりも順序が前であれば0未満の値（通常は-1）を、1つ目が2つ目よりも順序が後であれば0より大きい値（通常は1）を、1つ目と2つ目が順序は同じである、または、同値であると判定すれば0を戻り値とすること。

## サンプル

	Sub Test
		Dim Fruits As XArray
		Dim NumberOfFruits As Long
		Dim i As Long
		Set Fruits = New XArray
		Fruits.Add "Cherry"
		Fruits.Add "Apple"
		Fruits.Add "Banana"
		Fruits.Sort
		For i = 0 To Fruits.Count - 1
			Fruits.Item(i) = (i + 1) & ": " & Fruits.Item(i)
		Next
		MsgBox Fruits.Item(0) '1: Apple
		MsgBox Fruits.Item(1) '2: Banana
		MsgBox Fruits.Item(2) '3: Cherry
	End Sub
