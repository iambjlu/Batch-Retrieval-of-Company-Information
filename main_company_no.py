import requests
import xml.etree.ElementTree as ET
import csv

'''
使用說明:
把統編放進companies.txt(與此程式檔案置於同樣目錄)
可以獲得output.csv
'''

# 讀取統一編號列表
def read_accounting_numbers(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        numbers = [line.strip() for line in file if line.strip()]
    return numbers

# 從 API 獲取資料
def fetch_data_from_api(accounting_number):
    accounting_number_str = str(accounting_number)
    base_url = "https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=xml&$filter=Business_Accounting_NO%20eq%20" + accounting_number_str + "&$skip=0&$top=50"
    response = requests.get(base_url)
    if response.status_code == 200:
        return response.text
    else:
        print(f"Failed to fetch data for {accounting_number}. HTTP {response.status_code}")
        return None

# 解析 XML 資料
def parse_xml_data(xml_data):
    root = ET.fromstring(xml_data)
    rows = []
    for row in root.findall(".//row"):
        data = {
            "Business_Accounting_NO": row.findtext("Business_Accounting_NO", default=""),
            "Company_Name": row.findtext("Company_Name", default=""),
            "Company_Status_Desc": row.findtext("Company_Status_Desc", default=""),
            "Capital_Stock_Amount": row.findtext("Capital_Stock_Amount", default=""),
            "Paid_In_Capital_Amount": row.findtext("Paid_In_Capital_Amount", default=""),
            "Responsible_Name": row.findtext("Responsible_Name", default=""),
            "Register_Organization_Desc": row.findtext("Register_Organization_Desc", default=""),
            "Company_Location": row.findtext("Company_Location", default=""),
            "Company_Setup_Date": row.findtext("Company_Setup_Date", default=""),
            "Change_Of_Approval_Data": row.findtext("Change_Of_Approval_Data", default=""),
            "Revoke_App_Date": row.findtext("Revoke_App_Date", default=""),
            "Case_Status": row.findtext("Case_Status", default=""),
            "Case_Status_Desc": row.findtext("Case_Status_Desc", default=""),
            "Sus_App_Date": row.findtext("Sus_App_Date", default=""),
            "Sus_Beg_Date": row.findtext("Sus_Beg_Date", default=""),
            "Sus_End_Date": row.findtext("Sus_End_Date", default="")
        }
        rows.append(data)
    return rows

# 保存資料到 CSV
def save_to_csv(data, output_file):
    if not data:
        print("No data to save.")
        return

    fieldnames = list(data[0].keys())
    with open(output_file, 'w', encoding='utf-8-sig', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Data saved to {output_file}")

# 主程式
def main(input_file, output_file):
    accounting_numbers = read_accounting_numbers(input_file)
    all_data = []

    for number in accounting_numbers:
        print(f"Fetching data for {number}...")
        xml_data = fetch_data_from_api(number)
        if xml_data:
            parsed_data = parse_xml_data(xml_data)
            if parsed_data:  # 如果有解析到資料
                all_data.extend(parsed_data)
            else:  # 無資料時新增 "找不到" 記錄
                all_data.append({
                    "Business_Accounting_NO": number,
                    "Company_Name": "沒有可用的資料",
                    "Company_Status_Desc": "沒有可用的資料",
                    "Capital_Stock_Amount": "沒有可用的資料",
                    "Paid_In_Capital_Amount": "沒有可用的資料",
                    "Responsible_Name": "沒有可用的資料",
                    "Register_Organization_Desc": "沒有可用的資料",
                    "Company_Location": "沒有可用的資料",
                    "Company_Setup_Date": "沒有可用的資料",
                    "Change_Of_Approval_Data": "沒有可用的資料",
                    "Revoke_App_Date": "沒有可用的資料",
                    "Case_Status": "沒有可用的資料",
                    "Case_Status_Desc": "沒有可用的資料",
                    "Sus_App_Date": "沒有可用的資料",
                    "Sus_Beg_Date": "沒有可用的資料",
                    "Sus_End_Date": "沒有可用的資料"
                })
        else:
            # 無法取得資料時新增 "找不到" 記錄
            all_data.append({
                "Business_Accounting_NO": number,
                "Company_Name": "找不到資料",
                "Company_Status_Desc": "找不到資料",
                "Capital_Stock_Amount": "找不到資料",
                "Paid_In_Capital_Amount": "找不到資料",
                "Responsible_Name": "找不到資料",
                "Register_Organization_Desc": "找不到資料",
                "Company_Location": "找不到資料",
                "Company_Setup_Date": "找不到資料",
                "Change_Of_Approval_Data": "找不到資料",
                "Revoke_App_Date": "找不到資料",
                "Case_Status": "找不到資料",
                "Case_Status_Desc": "找不到資料",
                "Sus_App_Date": "找不到資料",
                "Sus_Beg_Date": "找不到資料",
                "Sus_End_Date": "找不到資料"
            })

    save_to_csv(all_data, output_file)

# 執行
if __name__ == "__main__":
    input_file = "companies.txt"  # 替換成你的 txt 文件路徑
    output_file = "output.csv"  # 輸出 CSV 文件名稱
    main(input_file, output_file)
