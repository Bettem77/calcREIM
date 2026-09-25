# calcREIM

粉体の複素有効質量(effective mass)を計算するプロジェクト。振動加振実験の測定データ(xlsx)から
複素有効質量 M と損失係数 η* を計算し、粒径別にまとめる。

## ファイル命名規則

計測データは `eta_star_v5_<label>_<f>Hz_<level>.xlsx` の形式で統一されている。

- `label`: 粒径ラベル。`<N>um`(例: `30um`, `75um`, `100um`, `200um`, `300um`)、または空ボックス測定
  を表す `box`。
- `f`: 加振周波数 [Hz](例: `50`, `100`, `200`, `300`, `500`, `1000`, `1500`)。
- `level`: 振動レベル。通常は `10`/`20`/`30`/`40`(だいたい ×10 で加速度A[m/s^2]に対応)。
  まれに `35` のような非標準値もあり、その場合は変換せずそのまま維持する。

例: `eta_star_v5_30um_500Hz_20.xlsx`、`eta_star_v5_box_1000Hz_10.xlsx`

## プロジェクト共通の注意点(全スクリプト)

- **Windowsパスのraw文字列化**: このプロジェクトの各スクリプトは冒頭で
  `INPUT_DIR = "C:\Users\USER\Desktop\calcREIM"` のように書かれがちだが、raw文字列にしないと
  `\U` がUnicodeエスケープと誤認識されて `SyntaxError: (unicode error) 'unicodeescape' codec...`
  になる。**新しいスクリプトを追加・コピーしたときは必ず `INPUT_DIR = r"C:\Users\USER\Desktop\calcREIM"`
  のように `r` プレフィックスを付けること**(`effective_mass_batch.py`、`force_waveform_batch.py`
  の両方で実際に発生した)。
- **依存パッケージ**: `pandas`, `numpy`, `openpyxl`, `matplotlib` が必要(`python -m pip install
  pandas numpy openpyxl matplotlib` で導入済み)。
- **PNG保存時の `OSError: [Errno 22] Invalid argument`**: コードの不具合ではなく、たいてい出力先の
  pngファイルを画像ビューアなど別プロセスで開いたままファイルロックされているのが原因。該当ファイルを
  閉じてから再実行すれば解消する。

### 旧命名規則からの変換ルール

過去に別の命名規則(`eta_star_v5_D<N>_<f>Hz_<level>.xlsx`、level が `1`/`2`/`3`/`4`)で追加された
ファイルが混入することがある。新しいファイルを追加したときにこの旧形式であれば、以下のルールで
統一後の命名にリネームする(このセッションで何度か発生した作業):

- `D<N>` → `<N>um`(例: `D30` → `30um`、`D100` → `100um`)
- level: `1`→`10`、`2`→`20`、`3`→`30`、`4`→`40`
- 上記以外の値(`35` など)はそのまま変更しない
- リネーム前に、変換後のファイル名が既存ファイルと衝突しないか必ず確認する

## effective_mass_batch.py

`eta_star_v5_*.xlsx` を一括読み込みし、複素有効質量 M と η* を計算して
`output/effective_mass_summary.xlsx` と `output/effective_mass_by_size.png` を出力するスクリプト。

- 入力: スクリプト内 `INPUT_DIR` で指定したフォルダ内の xlsx 群(現在はこのプロジェクトのルート)
- `EtaStar_Acc` シート(B6=f, B7=m_p, B8=m_c, B38=η* 検算用)と `TimeData` シート(時刻・力・加速度)
  を読む
- 出力先: `./output/`

## force_waveform_batch.py

`eta_star_v5_*.xlsx` の `TimeData`(A=時刻[s], B=力[N])から、計算・差し引き・規格化をせず力波形を
そのままグラフ化するスクリプト。粒径ラベルごとに `OUTPUT_DIR/waveform_<label>.png` を出力
(行=周波数, 列=加速度レベル)。出力先: `./output_waveform/`
