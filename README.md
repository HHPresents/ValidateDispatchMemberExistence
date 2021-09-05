ValidateDispatchMemberExistence
====

これは何？
----

COMオブジェクト(IDispatch)のメソッドおよびプロパティの存在確認をIDispatchのGetIDsOfNamesを使用して検証するサンプルプログラムです。

たとえば、Microsoft Office の Excel.Application や Word.Application は、バージョンによって使用できるメソッドやプロパティが異なります。そのため、バージョン情報によって使用できるかどうかを検証する必要がありますが、直接メソッドやプロパティの存在を確認した方が確実です。

技術的な説明
----

IDispatchにはGetIDsOfNameというメソッドがあります。このメソッドを使用すると、メンバー名(メソッドおよびプロパティの名前)からDISPIDを得ることができるのですが、メンバーが存在しなかった場合はDISP_E_UNKNOWNNAME(0x80020006)というエラーを返します。

このような特徴を利用して、IDispatch上の動的メンバーの存在有無を検証しようという魂胆です。

まず、C#にはIDispatchのインターフェース定義が無いみたいなので、GetIDsOfNameを呼び出すためだけのIDispatchのインターフェース定義をプライベートに用意します。

IDispatchには４種類のメソッドが定義されているのですが、今回使用するのはGetIDsOfNameだけです。そのため、GetIDsOfName以外のメソッドのシグネチャーはデタラメでも構いません。しかし、メンバーの定義順序はvtblの順序と直結するため省略してはいけません。
GetIDsOfNameは３番目に定義されていますので、GetIDsOfNameよりも前に宣言されているGetTypeInfoCountとGetTypeInfoの定義は必ず必要です。一方、InvokeはGetIDsOfNameよりも後に定義されているので省略しても構いません。(サンプルでは省略せずに書いてしまいました)

後は、検証したいCOMオブジェクトをIDispatchに型変換してGetIDsOfNameを呼び出してみるだけです。GetIDsOfNameにPreserveSig属性をつけていればHRESULT値が直接入手できるので、DISP_E_UNKNOWNNAME(0x80020006)の発生有無を直接検証できます。
※PreserveSig属性がついてないと .NET Framework によって COMException がスローされてしまいます。例外をキャッチしてHRESULT値を検証する方法でもよいのですが、処理が冗長になるのでお勧めしません。