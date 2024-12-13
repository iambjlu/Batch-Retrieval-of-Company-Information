import requests
import xml.etree.ElementTree as ET
import csv


'''
使用說明:
把公司名字關鍵字放進companies.txt(與此程式檔案置於同樣目錄)
可以獲得output.csv
'''


# 讀取公司名稱列表
def read_company_names(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        companies = [line.strip() for line in file if line.strip()]
    return companies


# 從 API 獲取資料
def fetch_data_from_api(company_name):
    base_url = "http://data.gcis.nat.gov.tw/od/data/api/6BBA2268-1367-4B42-9CCA-BC17499EBE8C"
    params = {
        "$format": "xml",
        "$filter": f"Company_Name like {company_name} and Company_Status eq 01",
        "$skip": 0,
        "$top": 50
    }
    response = requests.get(base_url, params=params)
    if response.status_code == 200:
        return response.text
    else:
        print(f"Failed to fetch data for {company_name}. HTTP {response.status_code}")
        return None


# 解析 XML 資料
def parse_xml_data(xml_data):
    root = ET.fromstring(xml_data)
    rows = []
    for row in root.findall(".//row"):
        data = {
            "Business_Accounting_NO": row.findtext("Business_Accounting_NO", default=""),
            "Company_Name": row.findtext("Company_Name", default=""),
            "Company_Status": row.findtext("Company_Status", default=""),
            "Company_Status_Desc": row.findtext("Company_Status_Desc", default=""),
            "Capital_Stock_Amount": row.findtext("Capital_Stock_Amount", default=""),
            "Paid_In_Capital_Amount": row.findtext("Paid_In_Capital_Amount", default=""),
            "Responsible_Name": row.findtext("Responsible_Name", default=""),
            "Register_Organization_Desc": row.findtext("Register_Organization_Desc", default=""),
            "Company_Location": row.findtext("Company_Location", default=""),
            "Company_Setup_Date": row.findtext("Company_Setup_Date", default=""),
            "Change_Of_Approval_Data": row.findtext("Change_Of_Approval_Data", default="")
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
    companies = read_company_names(input_file)
    all_data = []

    for company in companies:
        print(f"Fetching data for {company}...")
        xml_data = fetch_data_from_api(company)
        if xml_data:
            parsed_data = parse_xml_data(xml_data)
            all_data.extend(parsed_data)

    save_to_csv(all_data, output_file)


# 執行
if __name__ == "__main__":
    input_file = "companies.txt"  # 替換成你的 txt 文件路徑
    output_file = "output.csv"  # 輸出 CSV 文件名稱
    main(input_file, output_file)