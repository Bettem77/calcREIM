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

## フォルダ構成

ファイル数が増えたため、生データと計算用xlsxをサブフォルダに分けている(全スクリプトが対応済み)。

- `Rawdata/`: DSAの生データCSV(`<label>_<f>Hz_<level>_<date>.csv`)を置く場所。
  `calc_eta_star.py` の `RAWDATA_DIR` が参照する。
- `eta_star/`: `eta_star_v5_*.xlsx`(計算用ファイル本体、テンプレートの `eta_star_v5_box_1000Hz_10.xlsx`
  含む)を置く場所。`effective_mass_batch.py` / `force_waveform_batch.py` / `calc_eta_star.py` の
  `ETASTAR_DIR` が参照する。
- スクリプト本体(`*.py`)、`CLAUDE.md`、`output/`、`output_waveform/`、
  `データ取得状況一覧_ひな形.xlsx` はプロジェクトルート直下のまま。
- 新しいスクリプトを追加するときも、xlsxは `ETASTAR_DIR`、生データCSVは `RAWDATA_DIR` を参照するように
  すること(ルート直下を直接globしない)。

`calc_eta_star.py`(生データCSVからの自動生成分)は、同じ label/f/level の過去データと衝突しない
よう末尾に生データの日付を付けた `eta_star_v5_<label>_<f>Hz_<level>_<date>.xlsx`(例:
`eta_star_v5_100um_50Hz_10_0926.xlsx`)になる。`effective_mass_batch.py` /
`force_waveform_batch.py` のファイル名パターンは、この末尾日付があってもなくても解釈できるように
なっている(`level` の後ろの `_<date>` は任意)。

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
- **既存の `eta_star_v5_*.xlsx` を上書きしてはいけない**: これらは実測データそのものであり、
  スクリプトのバグ調査やテストのために安易に上書き・再生成しない。テストする場合は必ずプロジェクト外の
  一時フォルダ(例: `C:\Users\USER\Desktop\calcREIM\_drytest` のような作業用ディレクトリ、実行後に削除)
  にコピーしてから行うこと。`calc_eta_star.py` も既定では既存ファイルを上書きしない(後述)。

### 旧命名規則からの変換ルール

過去に別の命名規則(`eta_star_v5_D<N>_<f>Hz_<level>.xlsx`、level が `1`/`2`/`3`/`4`)で追加された
ファイルが混入することがある。新しいファイルを追加したときにこの旧形式であれば、以下のルールで
統一後の命名にリネームする(このセッションで何度か発生した作業):

- `D<N>` → `<N>um`(例: `D30` → `30um`、`D100` → `100um`)
- level: `1`→`10`、`2`→`20`、`3`→`30`、`4`→`40`
- 上記以外の値(`35` など)はそのまま変更しない
- リネーム前に、変換後のファイル名が既存ファイルと衝突しないか必ず確認する

## effective_mass_batch.py

`eta_star/` 内の `eta_star_v5_*.xlsx` を一括読み込みし、複素有効質量 M と η* を計算して
`output/effective_mass_summary.xlsx` と `output/effective_mass_by_size.png` を出力するスクリプト。

- 入力: `ETASTAR_DIR`(既定 `eta_star` フォルダ)内の xlsx 群
- `EtaStar_Acc` シート(B6=f, B7=m_p, B8=m_c, B38=η* 検算用)と `TimeData` シート(時刻・力・加速度)
  を読む
- 出力先: `./output/`
- 同じ条件(粒径・f・level)で日付違いなど複数ファイルがある場合、グラフは点ごとの単純プロットではなく
  同じfでの平均値±標準偏差のエラーバー表示になる(1点だけの条件は誤差0のマーカーになるだけ)。

## force_waveform_batch.py

`eta_star/` 内の `eta_star_v5_*.xlsx` の `TimeData`(A=時刻[s], B=力[N])から、計算・差し引き・規格化を
せず力波形をそのままグラフ化するスクリプト。粒径ラベルごとに `OUTPUT_DIR/waveform_<label>.png` を出力
(行=周波数, 列=加速度レベル)。入力: `ETASTAR_DIR`(既定 `eta_star` フォルダ)。出力先: `./output_waveform/`

## calc_eta_star.py

DSA(動的信号解析器)が出力する生データCSV(`[Header]`/`[Calibration]`/`[TIME_INST]` 形式)を、
`eta_star_v5_<label>_<f>Hz_<level>_<date>.xlsx` の `TimeData` シートに自動転記するスクリプト
(手作業の「A列=時刻・B列=力・D列=加速度を貼り付け」を代替する)。

- 入力: `RAWDATA_DIR`(既定 `Rawdata` フォルダ)内の `<label>_<f>Hz_<level>_<date>.csv`
  (例: `Rawdata/100um_50Hz_5_0926.csv`)。`[Calibration]` の `EU/V` 行から換算係数を読み、
  単位ラベルが `N`/`m/s2` のチャンネルをそれぞれ力・加速度として使う(チャンネル番号を決め打ちしない)。
- 出力先: `ETASTAR_DIR`(既定 `eta_star` フォルダ)。出力ファイル名は生データの日付をそのまま引き継ぐ
  (`eta_star_v5_100um_50Hz_10_0926.xlsx` など)ので、過去の同名(日付なし)ファイルと衝突しない。
- **力の符号に注意**: `EU/V × 電圧` をそのまま使うと、既存の正しいxlsx(FとaがF≈m·aで正相関)に対して
  符号が逆になることを実測で確認済み(相関 +0.97 → -0.98)。そのためスクリプト内で
  `FORCE_SIGN = -1` を掛けて補正している。センサー配線を変えない限りこの補正は変えない。
- **既定では既存ファイルを上書きしない**(出力先の `eta_star_v5_*.xlsx` が既に存在する場合はスキップ
  してメッセージを表示する)。あえて上書きしたい場合のみ `python calc_eta_star.py --overwrite` を使う。
- 実行すると対話的に以下を聞かれる:
  - 粉体質量 m_p [g](粒径labelごとに1回。空欄なら `EtaStar_Acc!B7` は未入力のまま出力)
  - 使用するデータ点数 N(周波数fごとに1回。先頭から何点使うか。空欄なら全点使用)
  容器質量 `B8` は既定値 `M_C_DEFAULT = 0.88` を仮入力する(対話入力なし)。

## eta_star_summary.py

`eta_star/` 内の `eta_star_v5_*.xlsx`(box除く)から η* / Re(M) / Im(M) を集計し、社内の既存資料と
同じ配列形式でまとめるスクリプト。η* は effective_mass_batch.py と同じ式で TimeData から都度
再計算する(シートのキャッシュ値には頼らない。空ボックス差し引き前の、そのファイル1本の値)。
同じ D・A(=level×10)・f の条件で日付違いなど複数ファイルがある場合は平均値を使う。

- 出力: `output/eta_star_summary.xlsx`
  - `D_f_table` シート: 行=(D[mm], A[m/s2])、列=f[Hz]、値=η* の一覧表
  - `by_frequency` シート: f[Hz]ごとに D/a_amp/η*/Re/Im の表を横に並べたもの
  - `raw` シート: 集計前の生データ(ファイルごとの1行、検算用)
- `effective_mass_batch.py` / `force_waveform_batch.py` / `eta_star_summary.py` は Excelで
  ファイルを開いているときにできるロックファイル(`~$eta_star_v5_....xlsx`)を除外して読む
  (`~$` で始まるファイル名は無視する)。
