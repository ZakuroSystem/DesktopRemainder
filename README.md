# デスクトップリマインダー Win32 API 版

C++ / Win32 API で実装した軽量版です。今回の版では、**常駐監視プロセス**と**編集GUIプロセス**を分離しています。

以下のパッケージを必ずインストールしてください。
https://www.microsoft.com/ja-jp/download/details.aspx?id=48145

## 構成

通常起動時の動きは次のとおりです。

- `DesktopReminder.exe` を通常起動すると、まず常駐用の **agent プロセス** を起動します
- その後、表示と編集を担当する **editor プロセス** を起動します
- editor を閉じても agent は残り、通知監視だけ継続します
- トレイアイコンから editor を再度開けます

### agent の役割

- タスクトレイ常駐
- 次回通知時刻の計算
- ポップアップ通知 / 通知領域通知
- CSV 再読込とスケジュール更新

### editor の役割

- タスク一覧表示
- 追加 / 編集 / 削除
- 完了 / 未完了
- 有効 / 無効
- スヌーズ
- CSV 読込 / 保存

## 主な機能

- 小窓表示
- タスク右クリックで追加 / 編集 / 削除 / 完了 / 未完了 / 有効 / 無効 / スヌーズ
- CSV 入出力
- 日本語 UI
- 年次 / 月次 / 日次タスク
- ポップアップ / 通知領域 / 両方 / なし の通知方式
- スヌーズ 5 / 10 / 30 分
- 元に戻す
- 初期タスクなし

## 軽量化のために入れていること

### 1. 常駐監視と編集GUIの分離

常駐時には editor の一覧や描画部品を保持しません。  
そのため、editor を閉じたあとのメモリ消費をさらに減らしやすくしています。

### 2. 毎分ポーリングをやめ、次回通知時刻まで待機

毎分全タスクを走査するのではなく、**次に通知が必要な最短時刻**まで待機します。  
これにより、待機中の CPU 使用率をほぼゼロに寄せやすくしています。

### 3. 文字列の使い回し

一覧表示で毎回予定文字列や通知文字列を再構築せず、タスク読込・編集時にキャッシュします。

### 4. /MD 設定

この版では CRT を **`/MD`** にしています。

#### 用語

- **CRT**: C Runtime Library。C/C++ の基本関数群やメモリ管理に使うランタイムです
- **`/MT`**: CRT を exe に静的リンクする設定
- **`/MD`**: CRT を DLL として共有利用する設定

#### `/MT` の特徴

- exe 単体で動きやすい
- 配布が楽
- その代わり exe がやや大きくなりやすい
- プロセスごとにランタイムを抱えやすい

#### `/MD` の特徴

- exe が小さくなりやすい
- ランタイム DLL を他プロセスと共有できる
- 専有メモリが少し下がる場合がある
- 配布先に **Visual C++ 再頒布可能パッケージ** が必要になることがある

つまり、今回の版は **単体配布最優先ではなく、軽さ寄り**の設定です。

## ビルド方法

Visual Studio の「x64 Native Tools Command Prompt for VS 2022」または「Developer Command Prompt for VS 2022」で、`DesktopReminder` フォルダへ移動して実行します。

```bat
msbuild DesktopReminder.vcxproj /p:Configuration=Release /p:Platform=x64
```

生成物:

```text
..\bin\Release\DesktopReminder.exe
```

## 実行方法

通常はそのまま起動してください。

```bat
DesktopReminder.exe
```

内部的には次の起動モードがあります。

- `DesktopReminder.exe --agent`
- `DesktopReminder.exe --editor`

通常は手動指定不要です。

## 補足

- 一般公開前の確認結果は `OPEN_SOURCE_READINESS.md` を参照してください
- 既定の CSV は `ドキュメント\tasks.csv` です
- セキュリティ上の理由で、CSV パスはローカルファイルのみ許可し、UNC などのネットワーク / デバイスパスは拒否します
- 設定ファイルは既定で `%LOCALAPPDATA%\DesktopReminder\DesktopReminder.ini` に保存され、現在の CSV パスや表示フィルタを保持します（旧版の exe 同フォルダ `DesktopReminder.ini` がある場合は初回起動時に移行を試みます）
- editor を閉じても agent は残ります
- 完全終了はトレイアイコン右クリックの **終了** です
