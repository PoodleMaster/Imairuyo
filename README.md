# Imairuyo（イマイルヨ）

## Imairuyo for PowerShell

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
| Restricted（制限付き） | すべてのスクリプトが実行不可（デフォルト設定）|
| RemoteSigned           | ローカルスクリプトは制限なく実行でき、外部からダウンロードしたスクリプトには署名が必要 |
| AllSigned              | 信頼された発行元から署名されたスクリプトのみが実行可能|
| Unrestricted           | すべてのスクリプトが実行可能ですが、外部からダウンロードしたスクリプトには警告が表示 |

### (2) 実行方法
アイコンを右クリックして「`PowerShell で実行`」を選択し起動してください。<BR>
起動するとPowerShellウィンドウが開きイマイルヨが実行されます。<BR>

![PowerShell で実行](https://github.com/user-attachments/assets/46b80683-294f-4b50-8d99-7c42613b61ec)

#### ❖ imairuyo.pas1実行例
![imairuyo.pas1](https://github.com/user-attachments/assets/81e67ede-fcbf-416d-ae81-685a63e259d7)

#### ❖ imairuyo_time.pas1実行例
![imairuyo_time.pas1](https://github.com/user-attachments/assets/97a58763-3baa-4c10-a2c3-b48a0b393993)

### (3) 終了方法
終了方法は、PowerShellウィンドウを閉じてください。<BR>

## Imairuyo for WindowsScriptHost（VisualBasic）


### (1) 実行方法

### (2) 終了方法

