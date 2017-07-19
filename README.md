xlsx ファイルのシートを分割して書き出すスクリプトです。

## 環境

#### Python の確認

Mac OS にプリインストールされている Python を使う

```
$ python --version
```

#### pip のインストール

```
$ pip -V

// pip がインストールされていない場合は、以下を実行する
$ curl -kL https://raw.github.com/pypa/pip/master/contrib/get-pip.py | python
```

#### openpyxl のインストール

```
$ pip install openpyxl
```

## 実行

- index.py と同じディレクトリに _in/ および _out/ ディレクトリを作成する

- _in/ ディレクトリに、シート分割したい xlsx ファイルを置く

- スクリプトを実行する

```
$ python index.py _in _out

// もしくは、以下でも可
$ ./start.sh
```

- _out/ ディレクトリに、分割された xlsx ファイルが出力される
