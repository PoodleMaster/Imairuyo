# Imairuyo（イマイルヨ）
**イマイルヨは、Teams等で勝手に退席中にならないようにするめのツールです。**<BR>
イマイルヨの[記事](https://qiita.com/PoodleMaster/items/4c2ad1501e615b164f7f)はこちらにあります。

## ■ Imairuyo for PowerShell

### (1) PowerShellスクリプトが実行できるか確認
PowerShellスクリプトが実行できるか確認するために、PowerShellを開いて「`Get-ExecutionPolicy`」を実行してください。<BR>
```
PS C:\Users\user> Get-ExecutionPolicy
```

実行ポリシー種別が、「`Restricted`」となっている場合は、PowerShellでのスクリプト実行ができません。<BR>
その場合は、**管理者権限**でPowerShellを開いて、「`Set-ExecutionPolicy RemoteSigned`」を実行し、実行ポリシーの変更をしてください。<BR>

```
PS C:\Users\user> Set-ExecutionPolicy RemoteSigned
```

| 実行ポリシー種別| 説明|
|------------------------|----------------------------------------------------------------------|
| Restricted（制限付き）  | すべてのスクリプトが実行不可（デフォルト設定）|
| RemoteSigned           | ローカルスクリプトは制限なく実行でき、外部からダウンロードしたスクリプトには署名が必要 |
| AllSigned              | 信頼された発行元から署名されたスクリプトのみが実行可能|
| Unrestricted           | すべてのスクリプトが実行可能ですが、外部からダウンロードしたスクリプトには警告が表示 |

> プログラムをポリシー指定することで実行できる場合もあります。<BR>
> 以下のようにバッチファイルを作っておくと便利です。<BR>
> バッチファイル ： imairuyo.bat<BR>
> ```bat:imairuyo.bat
> powershell -ExecutionPolicy RemoteSigned -File imairuyo.ps1
> ```

### (2) 実行方法
「`imairuyo.ps1`」または「`imairuyo_time.ps1`」アイコンを右クリックして「`PowerShell で実行`」を選択し起動してください。<BR>
起動するとPowerShellウィンドウが開きイマイルヨが実行されます。<BR>

<img width="400" src="https://github.com/user-attachments/assets/46b80683-294f-4b50-8d99-7c42613b61ec">

#### ❖ imairuyo.ps1実行例
![imairuyo.pas1](https://github.com/user-attachments/assets/81e67ede-fcbf-416d-ae81-685a63e259d7)

#### ❖ imairuyo_time.ps1実行例
![imairuyo_time.pas1](https://github.com/user-attachments/assets/97a58763-3baa-4c10-a2c3-b48a0b393993)

#### ❖ imairuyo.bat
こちらはバッチファイルなので、ダブルクリックで起動可能です。

### (3) 終了方法
終了方法は、起動時にできたPowerShellウィンドウを閉じてください。<BR>

## ■ Imairuyo for WindowsScriptHost（VisualBasic）
PowerShellでうまくいかない方は、こちらをお試しください。

### (1) 実行方法
#### ❖ imairuyo_start.vbs実行例<BR>
「`imairuyo_start.vbs`」アイコンを右クリックして「`プログラムから開く` ➜ `Microsoft®Windows Based Script Host`」を選択し起動してください。
起動するとイマイルヨが実行されます。（**ウィンドウなどは表示されません**）<BR>

<img width="600" src="https://github.com/user-attachments/assets/586bea07-2698-4cbe-befa-efc380d23dcb">

タスクマネージャー上では、以下のようにアプリが動作していることが分かります。
![タスクマネージャー](https://github.com/user-attachments/assets/a08885a9-34b2-40fd-8c4b-e468677d0dff)

### (2) 終了方法
#### ❖ imairuyo_stop.vbs実行例<BR>
「`imairuyo_stop.vbs`」アイコンを右クリックして「`プログラムから開く` ➜ `Microsoft®Windows Based Script Host`」を選択し起動してください。
起動するとイマイルヨが実行されます。<BR>

<img width="600" src="https://github.com/user-attachments/assets/9bb1cddb-2ea3-4b81-be40-ad0495943573">

タスクマネージャー上から終了する場合は、以下のようにアプリを指定して終了させます。

### (3) 手動終了方法
基本的に終了は、「`imairuyo_stop.vbs`」で問題ありません。
何か問題があった場合に備えて、手動終了手順を記載しておきます。

#### ❖ プロセスID確認方法
イマイルヨのプロセスID確認方法は、コマンドプロンプトから以下のコマンドを実行してください。
```
> cd %TEMP%
> type imairuyo_start.log
imairuyo_start: 34408
```
上記実行例の場合、「34408」が、起動したイマイルヨの「プロセスID」となります。

#### ❖ タスクマネージャーからの終了方法
イマイルヨのプロセスIDを確認後、「タスクバー右クリック」➔「タスクマネージャー」を起動し、タスクマネージャーの「詳細」を押下し、検索バーに「script」等を入力してください。
![タスクマネージャー](https://github.com/user-attachments/assets/386ecc5a-6e11-48cd-ac53-f5ea71e0a1a1)
先ほど探したプロセスID（imairuyo_start: 34408）と同じプロセスIDがあるので、これがイマイルヨのプロセスになりますので、これを右クリックし「タスクの終了」を選択し、プロセスを終了させます。

#### ❖ ロックファイル手動削除方法
コマンドプロンプトからのロックファイルの確認方法、及び、手動削除は方法は、以下の通りです。<BR>
```
> cd %TEMP%
> dir imairuyo_lock.lck
> DEL imairuyo_lock.lck
```
> ※「`%TEMP%`」とは、「`C:\Users\Username\AppData\Local\Temp`」のことです。<BR>
> ※ PowerShellの場合、該当ディレクトリへの移動は「`cd $env:TEMP`」としてください。<BR>

#### ❖ 注意事項
※タスクマネージャーから「`タスクの終了`」を選択した場合は、ロックファイル（`imairuyo_lock.lck`）が残るため、「`imairuyo_start.vbs`」の新規起動ができなくなってしまいます。そのため、必ずロックファイルの削除をお願いいたします。ロックファイルは手動で削除する他、「`imairuyo_stop.vbs`」または「`imairuyo_recover.vbs`」を実行しても、ロックファイルの削除が可能です。

### (4) トラブルシューティング
「`imairuyo_recover.vbs`」アイコンを右クリックして「`PowerShell で実行`」を選択し起動してください。<BR>
処理としては以下3つの処理を実行します。

1. WScript.exe プロセスをすべて終了（「`imairuyo_recover.vbs`」自身のプロセスは除外）
2. ロックファイルを削除
3. 「`imairuyo_recover.vbs`」自身のプロセスを終了

