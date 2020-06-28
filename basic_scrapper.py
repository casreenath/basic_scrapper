from scrapper import Scrapper
import sys
import time

try:
    scrapper_obj = Scrapper()
    print("Created OBJ")
    scrapper_obj.go_to_page(1)
    driver = scrapper_obj.get_driver()
except Exception as errmsg:
    print("Error: {}".format(errmsg))
    sys.exit(1)
company_data_list = []
count = 0
# =================================================
# Custom code to extract data
# =================================================
table_data = driver.find_elements_by_class_name("zp_3UsOq") # for the entire row data
for row_data in table_data:
    count += 1
    print("Count : {}".format(count))
    company_data_dict = {
        "Company_Name": "",
        "Employee_Headcount": "",
        "Industry_Sector": "",
        "Linkedin_URL": "",
        "FB": "",
        "Twitter": "",
        "Website": ""
    }
    name_data = row_data.find_elements_by_class_name("zp_BpGFW") # Title data
    company_data_dict['Company_Name'] = name_data[0].text
    # print(company_data_dict['Company_Name'])
    company_size_data = row_data.find_elements_by_class_name("zp_3jRyV") # Company Size
    company_data_dict['Employee_Headcount'] = company_size_data[0].text
    # print(company_data_dict['Employee_Headcount'])
    industry_sector_data = row_data.find_elements_by_class_name("zp_P8_of.zp_3YVzq")
    company_data_dict['Industry_Sector'] = industry_sector_data[0].text
    # print(company_data_dict['Industry_Sector'])
    social_urls = row_data.find_elements_by_class_name("zp_3_fnL")
    # print("Number of social_urls : {}".format(len(social_urls)))
    for social_url in social_urls:
        link = social_url.get_attribute('href')
        # print(link)
        if "linkedin" in link:
            company_data_dict['Linkedin_URL'] = link
        elif "facebook" in link:
            company_data_dict['FB'] = link
        elif "twitter" in link:
            company_data_dict['Twitter'] = link
        elif social_urls.index(social_url) == 0:
            company_data_dict['Website'] = social_urls[0].get_attribute('href')
    # zp_1IwK9 zp_3kWQr zp_2b_XK
    detailed_link_data = row_data.find_elements_by_class_name("zp_1IwK9.zp_3kWQr.zp_2b_XK")
    detailed_link = detailed_link_data[0].get_attribute('href')
    # print("detailed_link : {}".format(detailed_link))
    driver.execute_script("window.open('{}');".format(detailed_link))
    time.sleep(5)
    driver.switch_to.window(driver.window_handles[-1])
    detailed_datas = driver.find_elements_by_class_name("zp_2-WNE")
    revenue = ""
    location = ""
    for each in detailed_datas:
        detailed_data = each.text
        if "Annual Revenue" in detailed_data:
            revenue = detailed_data.replace("Annual Revenue", "")
            revenue = revenue.replace(" ", "")
        elif "india" in detailed_data.lower():
            location = detailed_data
    company_data_dict['Revenue'] = revenue.split("\n")[0]
    company_data_dict['Location'] = location.replace("\n", ",")
    print(company_data_dict)
    company_data_list.append(company_data_dict)
    driver.close()
    time.sleep(5)
    driver.switch_to.window(driver.window_handles[0])

# ===========================================================================

scrapper_obj.convert_to_excel(company_data_list, './scapper_output.xls')
