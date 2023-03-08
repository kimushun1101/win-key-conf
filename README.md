# win-key-conf

Windows のキーボード操作を効率化するプログラムです。
具体的には、AutoHotKey とGoogle 日本語入力、Ctrl2Cap の3 つのソフトウェアを用いて、日本語キーボードにある`無変換キー`と`変換キー`、`CapsLock キー`に異なる機能を与えます。
3 つのソフトはそれぞれ独立しているため、すべてを導入する必要はありません。
個人の好みに合ったソフトだけ使用してください。

## 導入方法
- お手元にGit 環境がある人：  
  クローンしてください。  
  `$ git clone git@github.com:kimushun1101/win-key-conf.git`
- Git 環境がない人：
  1. GitHub ページの`Code<>▼` から`Download ZIP`
  2. `win-key-conf-main.zip` をお好みのフォルダに展開してください。

各ソフトウェアの導入方法については以下をご覧ください。

---

## 無変換スイッチ
`無変換キー`と他のキーを同時押しすることで様々な動作ができるようになります。
詳細は
https://github.com/kimushun1101/muhenkan-switch
を御覧ください

### 導入方法
1. `install_muhenkan-switch.cmd` を実行
2. `muhenkan-switch\muhenkan.exe` を実行

## Google 日本語入力
日本語入力をサポートするソフトウェアです。
`GoogleIME\henkan_muhenkan.txt` は、私が使っている設定をGoogle 日本語入力でエクスポートしたものです。
### できること
- 設定する機能
  - `無変換キー`でIME オフ
  - `変換キー`でIME をオン
    - 入力中に`無変換キー`で半角に変換
    - 変換中に`BackSpace` で変換前に戻る
- Google 日本語入力自体の機能
  - (重要)以上の設定をテキストファイルにエクスポート or インポート
  - 方向キーが簡単に出せる
    - （日本語入力で）`zh` : ←
    - `zj` : ↓
    - `zk` : ↑
    - `zl` : →
  - 他にも様々な機能がありますので、ご興味のある方は調べてみてください。
### 導入方法
1. https://www.google.co.jp/ime/
ここからダウンロードしてインストール
2. タスクバーの時刻付近にある(デフォルトのMS IME を使用していた場合)`J`のアイコンをクリックしてGoogle 日本語入力に切り替え
3. `A`または`あ` のアイコンを右クリックしてプロパティをクリック
4. `キーの設定の選択`→`編集…`をクリック
5. `編集▼`から`インポート`をクリック
6. `GoogleIME\henkan_muhenkan.txt` を選択

## Ctrl2Cap
`CapsLock キー`を`Ctrl キー`に置き換えることができるソフトウェアです。
本当はAutoHotKey スクリプトで実現したかったのですが、日本語キーボードでは難しかったのでこちらで設定しました。
解決できる方がおりましたらPull Request、もしくはSNS などで教えてください。
### 導入方法
1. https://learn.microsoft.com/ja-jp/sysinternals/downloads/ctrl2cap
公式のホームページで内容を確認
2. `Ctrl2Cap\install_Ctrl2Cap.cmd` を右クリック→「管理者として実行」
3. Ctrl2cap successfully installed. You must reboot for it to take effect. と出ていたら再起動

---

## 設定を戻す・アンイストール
お好みの状態まで段階的に戻せます。
### AutoHotKey
1. `muhenkan-switch\uninstall.exe` を実行
2. `muhenkan-switch` を削除
### Google 日本語入力
1. キー設定をMicrosoft IME に戻す：導入方法5. の`編集▼`の`定義済みのキーマップからインポート`から`MS-IME`をクリック
2. Microsoft IME 自体に戻す： タスクバーの時刻付近にある青い丸のアイコンをクリックしてMicrosoft IME に切り替え
3. Google 日本語入力自体のアンイストール：Windows の設定→アプリと機能からGoogle 日本語入力を選択してアンイストール
### Ctrl2Cap
1. `Ctrl2Cap\uninstall_Ctrl2Cap.cmd` を右クリック→「管理者として実行」
2. Ctrl2cap uninstalled. You must reboot for this to take effect. と出ていたら再起動

---

## ライセンス
The MIT License