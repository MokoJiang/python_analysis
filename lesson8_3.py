import pandas as pd

def main():
    df = pd.read_csv('上市公司資料.csv')
    df1 = df.dropna()
    df2 = df1.reindex(columns=['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收'])
    df3 = df2.rename(columns={
        '營業收入-上月營收':'上月營收',
        '營業收入-當月營收':'當月營收'
        })
    df3.to_csv('上市公司資料整理.csv',encoding='utf-8')
    df3.to_excel('上市公司資料整理.xlsx')
    print("存檔完成")

if __name__ == '__main__':
    main()


import pandas as pd  # 導入 pandas 函式庫，用於數據處理
import json # 導入 json 函式庫，用於處理 json 數據
import os # 導入 os 函式庫，用於處理檔案路徑

def main():
    # 取得程式碼所在目錄的絕對路徑
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 建立上市公司資料CSV檔案的路徑
    csv_file_path = os.path.join(script_dir, '上市公司資料.csv')
    # 建立輸出CSV檔案的路徑
    output_csv_file_path = os.path.join(script_dir, '上市公司資料整理.csv')
    # 建立輸出Excel檔案的路徑
    output_excel_file_path = os.path.join(script_dir, '上市公司資料整理.xlsx')

    try:
        # 讀取上市公司資料CSV檔案，指定編碼為 UTF-8 以處理中文
        df = pd.read_csv(csv_file_path, encoding='utf-8') 
    except FileNotFoundError:
        # 處理檔案不存在的錯誤
        print(f"錯誤：找不到檔案 '{csv_file_path}'。請確認檔案存在於與程式碼相同的目錄下。")
        return  # 程式結束
    except pd.errors.ParserError:
        # 處理CSV解析錯誤
        print(f"錯誤：無法解析檔案 '{csv_file_path}'。請檢查檔案格式和編碼。")
        return # 程式結束
    except Exception as e:
        # 處理其他可能發生的錯誤
        print(f"讀取 CSV 檔案 '{csv_file_path}' 時發生意外錯誤：{e}")
        return # 程式結束


    try:
        # 讀取 YouBike 數據的 JSON 檔案
        with open(os.path.join(script_dir, 'youbike_data.json'), 'r', encoding='utf-8') as f:
            youbike_data = json.load(f)
        # 將 YouBike 數據轉換成 Pandas DataFrame
        youbike_df = pd.DataFrame(youbike_data)
    except FileNotFoundError:
        print(f"錯誤: 找不到檔案 'youbike_data.json'。請確認檔案存在於與程式碼相同的目錄下。")
        return
    except json.JSONDecodeError:
        print(f"錯誤: 'youbike_data.json' 檔案格式錯誤。請檢查檔案內容。")
        return
    except Exception as e:
        print(f"讀取或解析 'youbike_data.json' 檔案時發生錯誤: {e}")
        return


    df1 = df.dropna()  # 移除 DataFrame 中含有任何缺失值的列
    
    # 檢查必要的欄位是否存在，避免 KeyError 錯誤
    required_columns = ['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收']
    missing_columns = [col for col in required_columns if col not in df1.columns]
    if missing_columns:
        print(f"錯誤：'上市公司資料.csv' 檔案缺少以下欄位：{missing_columns}")
        return

    df2 = df1.reindex(columns=required_columns)  # 重新索引 DataFrame，只保留指定的欄位
    df3 = df2.rename(columns={  # 重新命名欄位
        '營業收入-上月營收':'上月營收',  # 將 '營業收入-上月營收' 重新命名為 '上月營收'
        '營業收入-當月營收':'當月營收'  # 將 '營業收入-當月營收' 重新命名為 '當月營收'
    })

    try:
        # 將處理後的數據保存到 CSV 檔案，並忽略索引
        df3.to_csv(output_csv_file_path, encoding='utf-8', index=False) 
        # 將處理後的數據保存到 Excel 檔案，並忽略索引
        df3.to_excel(output_excel_file_path, index=False) 
    except Exception as e:
        # 處理檔案保存錯誤
        print(f"保存檔案時發生錯誤：{e}")
        return

    print("檔案已成功保存")  # 顯示保存成功的訊息


if __name__ == '__main__':
    main()  # 執行主函數
