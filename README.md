# csv_viewer

巨大CSVを軽快に閲覧するための軽量ビューワ（Python / Tkinter製）。

## 特徴

- **行番号・列ヘッダ表示** とExcelライクな **矩形セル選択**（ドラッグ）
- **2段ヘッダ**: ラベルのドット前（接頭語）を共通列でまとめて上段に、ドット後（個別名）を下段に表示
  - 例: `FcmispMidsocOutParamArray.awbposgate_upos_cmn` → 上段 `FcmispMidsocOutParamArray` / 下段 `awbposgate_upos_cmn`
- **巨大CSV対応**: 行オフセットをバックグラウンド索引し、表示中の行だけ読む仮想スクロール（全行・全列をメモリに載せない）
- **先頭行・先頭列を固定**
- **検索**: 部分一致・大小無視・空白区切りの複数キーワード（AND）。セル値だけでなく **列ラベル名** も対象
- **キーボード操作**: 矢印キーで選択移動、PageUp/PageDown、Ctrl+C でコピー
- **ドラッグ&ドロップ** で開く（Windows）

---

## Windowsでの使い方

### 0. 前提: Python 3 のインストール

[python.org](https://www.python.org/downloads/) から Python 3 をインストールしてください。
インストーラ最初の画面で **「Add python.exe to PATH」にチェック**を入れます（`py` ランチャーが入ります）。

> tkinter（GUI）は Python に同梱されているので追加インストール不要です。

### 1. リポジトリを取得（OneDrive外のローカルへ）

> ⚠️ **OneDrive配下に置かないでください。** OneDrive が `.git` 内部を同期して破損・競合の原因になります。`C:\tools` などローカルディスクへ clone してください。

```cmd
cd /d C:\tools
git clone https://github.com/kazuhiko-kobayashi-dnjp/csv_viewer.git
cd csv_viewer
```

Git を使わない場合は、GitHub の「Code → Download ZIP」で取得し、`C:\tools\csv_viewer` などに解凍してもOKです。

### 2. 起動

`csv_viewer.vbs` を **ダブルクリック**するとビューワが開きます（コンソール窓は出ません）。

- 空の状態で起動 → ツールバーの **Open CSV** で選択、またはウィンドウへCSVをドラッグ&ドロップ
- `csv_viewer.vbs` のアイコンに **CSVファイルをドロップ** → そのファイルを直接開く

### 3. （任意）.csv の関連付け＋ドラッグ&ドロップ強化

PowerShell で以下を一度だけ実行すると、
`tkinterdnd2`（ウィンドウ内D&D用）の導入と、`.csv` ダブルクリックで本ツールが開く関連付けを行います。

```powershell
powershell -ExecutionPolicy Bypass -File install_windows.ps1
```

実行後は **CSVファイルをダブルクリック**するだけで開けます。

### うまく動かないとき

`csv_viewer.vbs` が即閉じる・エラーになる場合は、**`run_debug.bat`** をダブルクリックしてください。
コンソールにエラー内容（トレースバック）が表示され、ウィンドウが残ります。

Python環境の切り分けには **`diagnose.bat`** を使うと、どの Python が使われているか・正常に動くかを確認できます。

> 既知の注意点: PATH 上の `python` が壊れた古いインストールを指している環境でも、本ツールは Windows標準の **`py` ランチャー** 経由で起動するため影響を受けません。

---

## 操作方法

| 操作 | 動作 |
|------|------|
| ドラッグ（セル上） | 矩形選択 |
| 矢印キー | 選択セルを上下左右に移動 |
| PageUp / PageDown | ページ送り |
| Ctrl+C | 選択範囲をタブ区切りでコピー |
| 検索ボックスに入力 + Enter | 次を検索（Shift+Enter / ▲ で前を検索） |
| マウスホイール | 縦スクロール（Shift+ホイールで横スクロール） |

### 検索について

- **部分一致・大小無視**。前後の空白は無視されます。
- **複数キーワード**（空白区切り）は、同一行のどこかに全部あればヒット（AND検索）。
- **列ラベル名も検索対象**。長い列名の一部を入力すると、その列へジャンプします。

---

## WSL / Linux での実行（開発者向け）

```bash
sudo apt install python3-tk        # 未導入の場合
python3 csv_viewer.py [file.csv]
```

> WSLg のウィンドウは Windows エクスプローラーからのドラッグ&ドロップを受け取れません。
> エクスプローラーからのD&Dや `.csv` 関連付けを使いたい場合は、上記「Windowsでの使い方」のとおり **Windowsネイティブ** で実行してください。

---

## ファイル構成

| ファイル | 役割 |
|----------|------|
| `csv_viewer.py` | 本体（Windows / WSL 共通） |
| `csv_viewer.vbs` | Windows用ランチャー（コンソール窓なし、`py` ランチャー経由） |
| `run_debug.bat` | コンソール付きで起動しエラーを確認 |
| `diagnose.bat` | Python環境の診断 |
| `install_windows.ps1` | `tkinterdnd2` 導入 + `.csv` 関連付け |
| `requirements.txt` | 依存（任意の `tkinterdnd2`） |
