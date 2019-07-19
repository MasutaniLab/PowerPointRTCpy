# PowerPointRTCpy
PowerPointを操作するRTC Python版（升谷研版）

## 概要
- [オフィスソフトを操作するためのRTC群](https://www.openrtm.org/openrtm/ja/project/contest2014_6) に含まれるPowerPointRTCpyをOpenRTM-aist-1.2.0 + Python 3.7の環境で動くようにしたもの．

## 準備
- [Python for Windows (pywin32) Extensions](https://github.com/mhammond/pywin32)の[Release](https://github.com/mhammond/pywin32/releases)から`pywin32-224.win-amd64-py3.7.exe`をダウンロードし，実行してインストールする．
- https://github.com/MasutaniLab/PowerPointRTCpy からGitでクローン

## 仕様
http://officertcpy.kurushiunai.jp/15.html から引用．

### 入力ポート
|名称|データ型|説明|
|---|---|---|
|SlideNumberIn|TimedShort|コンフィギュレーションパラメータSlideNumberInRelativeが0のときは最初のスライドからの番号、1のときは現在のスライド番号からの番号|
|EffectNumberIn|TimedShort|実行するアニメーションの数|
|Pen|TimedShortSeq|描画する線の座標|

### 出力ポート
|名称|データ型|説明|
|---|---|---|
|SlideNumberOut|TimedShort|現在のスライド番号|

### コンフィギュレーション
|名称|型|デフォルト値|説明|
|---|---|---|---|
|SlideFileInitialNumber|int|0|スライドショー開始時のスライド番号|
|SlideNumberInRelative|int|1|0のときはSlideNumberInが最初のスライドからの番号になり、1のときは現在のスライド番号からの番号になる。|
|file_path|string|NewFile|開くPowerPointファイルの名前。"NewFile"と入力すると新規作成。|
