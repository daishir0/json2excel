import json
import pandas as pd
import sys
from typing import Dict, List, Any
import re

def flatten_json(json_obj: Dict[str, Any], parent_key: str = '', sep: str = '-') -> Dict[str, Any]:
    """
    ネストされたJSONオブジェクトをフラットな構造に変換する
    例: {'websites': {'official': 'url'}} → {'websites-official': 'url'}
    """
    items: List[tuple] = []
    for k, v in json_obj.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_json(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)

def extract_json_blocks(content: str) -> List[str]:
    """
    テキストからJSONブロックを抽出する
    """
    # 改行を統一
    content = content.replace('\r\n', '\n')
    
    # 空行を削除
    content = re.sub(r'\n\s*\n', '\n', content)
    
    # ```json タグを削除
    content = content.replace('```json\n', '')
    content = content.replace('```\n', '')
    content = content.replace('```', '')
    
    # 各JSONブロックを抽出
    blocks = []
    current_block = ''
    brace_count = 0
    
    for char in content:
        if char == '{':
            if brace_count == 0:
                current_block = '{'
            else:
                current_block += char
            brace_count += 1
        elif char == '}':
            brace_count -= 1
            if brace_count == 0:
                current_block += '}'
                if current_block.strip():
                    blocks.append(current_block.strip())
                current_block = ''
            else:
                current_block += char
        elif brace_count > 0:
            current_block += char
    
    return blocks

def process_json_file(file_path: str) -> pd.DataFrame:
    """
    JSONファイルを読み込み、DataFrameに変換する
    """
    print(f"Processing file: {file_path}")
    
    # データを格納するリスト
    data = []
    # 全てのキーを記録する集合
    all_keys = set()
    # 処理したレコード数
    processed_count = 0
    # エラーが発生したレコード数
    error_count = 0
    
    # ファイルを読み込む
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # JSONブロックを抽出
    json_blocks = extract_json_blocks(content)
    total_blocks = len(json_blocks)
    print(f"Found {total_blocks} JSON blocks")
    
    for block in json_blocks:
        try:
            # JSONをパース
            json_data = json.loads(block)
            # JSONをフラット化
            flat_data = flatten_json(json_data)
            # 全てのキーを記録
            all_keys.update(flat_data.keys())
            # データを追加
            data.append(flat_data)
            processed_count += 1
            if processed_count % 100 == 0:
                print(f"Processed {processed_count}/{total_blocks} records...")
        except json.JSONDecodeError as e:
            error_count += 1
            print(f"Error parsing JSON block: {str(e)[:100]}...")
            continue
    
    # 処理結果の表示
    print(f"\nProcessing completed:")
    print(f"Total blocks found: {total_blocks}")
    print(f"Successfully processed: {processed_count}")
    print(f"Failed to process: {error_count}")
    
    if not data:
        raise ValueError("No data was successfully processed")
    
    # DataFrameを作成
    df = pd.DataFrame(data)
    
    # 全てのカラムが存在することを確認（欠損値はNaN）
    for key in all_keys:
        if key not in df.columns:
            df[key] = None
    
    print(f"\nTotal columns: {len(df.columns)}")
    print("\nColumns found:")
    for col in sorted(df.columns):
        print(f"- {col}")
    
    return df

def main():
    if len(sys.argv) != 2:
        print("Usage: python json2excel.py <input_json_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = input_file.rsplit('.', 1)[0] + '.xlsx'
    
    print("Starting conversion process...")
    try:
        df = process_json_file(input_file)
        
        # Excelファイルに保存
        print(f"\nSaving to Excel file: {output_file}")
        df.to_excel(output_file, index=False)
        print(f"Conversion completed. {len(df)} records saved to {output_file}")
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()