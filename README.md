# json2excel

## Overview
json2excel is a Python script that converts nested JSON data into a flat Excel file. It's particularly useful when dealing with complex JSON structures that need to be analyzed in spreadsheet format. The script automatically flattens nested JSON objects using hyphen-separated keys and handles multiple JSON blocks within a single file.

## Installation
1. Clone the repository:
```bash
git clone https://github.com/daishir0/json2excel
```

2. Navigate to the project directory:
```bash
cd json2excel
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage
Run the script from the command line by providing the input JSON file path:

```bash
python json2excel.py <input_json_file>
```

The script will automatically create an Excel file with the same name as the input file but with a .xlsx extension.

### Input Example 1:
```json
{
    "name": "John Doe",
    "contact": {
        "email": "john@example.com",
        "phone": {
            "home": "123-456-7890",
            "work": "098-765-4321"
        }
    }
}
```

### Output Example 1:
Excel file with columns:
- name
- contact-email
- contact-phone-home
- contact-phone-work

### Input Example 2:
```json
{
    "product": "Laptop",
    "specifications": {
        "cpu": "Intel i7",
        "memory": {
            "size": "16GB",
            "type": "DDR4"
        }
    },
    "price": 1200
}
```

### Output Example 2:
Excel file with columns:
- product
- specifications-cpu
- specifications-memory-size
- specifications-memory-type
- price

## Notes
- The script automatically flattens nested JSON structures using hyphens (-) as separators
- Multiple JSON objects in the input file are processed as separate rows in the output Excel file
- Progress information and any parsing errors are displayed during processing
- The script handles JSON blocks that may be wrapped in Markdown-style code blocks (```json)

## License
This project is licensed under the MIT License - see the LICENSE file for details.

---

# json2excel

## 概要
json2excelは、ネストされたJSONデータをフラットなExcelファイルに変換するPythonスクリプトです。複雑なJSON構造をスプレッドシート形式で分析する必要がある場合に特に便利です。スクリプトは、ネストされたJSONオブジェクトをハイフン区切りのキーを使用して自動的にフラット化し、単一ファイル内の複数のJSONブロックを処理します。

## インストール方法
1. リポジトリをクローンします：
```bash
git clone https://github.com/daishir0/json2excel
```

2. プロジェクトディレクトリに移動します：
```bash
cd json2excel
```

3. 必要なパッケージをインストールします：
```bash
pip install -r requirements.txt
```

## 使い方
コマンドラインから入力JSONファイルのパスを指定してスクリプトを実行します：

```bash
python json2excel.py <input_json_file>
```

スクリプトは自動的に入力ファイルと同じ名前で.xlsx拡張子のExcelファイルを作成します。

### 入力例1：
```json
{
    "name": "John Doe",
    "contact": {
        "email": "john@example.com",
        "phone": {
            "home": "123-456-7890",
            "work": "098-765-4321"
        }
    }
}
```

### 出力例1：
以下のカラムを持つExcelファイル：
- name
- contact-email
- contact-phone-home
- contact-phone-work

### 入力例2：
```json
{
    "product": "Laptop",
    "specifications": {
        "cpu": "Intel i7",
        "memory": {
            "size": "16GB",
            "type": "DDR4"
        }
    },
    "price": 1200
}
```

### 出力例2：
以下のカラムを持つExcelファイル：
- product
- specifications-cpu
- specifications-memory-size
- specifications-memory-type
- price

## 注意点
- スクリプトは、ハイフン（-）をセパレータとして使用し、ネストされたJSON構造を自動的にフラット化します
- 入力ファイル内の複数のJSONオブジェクトは、出力Excelファイルの個別の行として処理されます
- 処理中の進捗情報やパースエラーが表示されます
- Markdown形式のコードブロック（```json）でラップされたJSONブロックも処理できます

## ライセンス
このプロジェクトはMITライセンスの下でライセンスされています。詳細はLICENSEファイルを参照してください。