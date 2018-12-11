import urllib3
import time
import re
import random
import certifi
from bs4 import BeautifulSoup
from openpyxl import load_workbook

http = urllib3.PoolManager(cert_reqs='CERT_REQUIRED', ca_certs=certifi.where())


def main():
    start = 0
    year_low = 2007
    year_high = 2018
    #   This is our year combination
    #   2007 - 2008 => Index 0
    #   2009 - 2010 => Index 1
    #   2011 - 2012 => Index 2
    #   2013 - 2014 => Index 3
    #   2015 - 2016 => Index 4
    #   2017 - 2018 => Index 5
    for i in range(0, 6):
        start = 0
        for j in range(0, 990):
            delay = random.randrange(3, 15)
            time.sleep(delay)
            page_scrape = 'https://scholar.google.ca/scholar?start=' + str(
                start) + '&hl=en&as_sdt=0,5&sciodt=0,5&as_ylo=' + str(year_low) + '&as_yhi=' + str(
                year_low + 1) + '&cites=10992335886072847329&scipsc='
            r = get_request(page_scrape)
            citation_papers = process_request(r)
            length = len(citation_papers)
            if length < 1:
                r = get_request(page_scrape)
                citation_papers = process_request(r)
                length = len(citation_papers)
                if length < 1:
                    print("::::::::::::::::::::DONE: Year Processed::::::::::::::::::::\n")
                    break
            else:
                print("Processing Tab: " + str(start) +
                      "\nYear Range: " + str(year_low) + " - " + str(year_low + 1) + "\n")
                process_data(citation_papers, start, year_low, year_low + 1)
                start = start + 10
        if year_low < year_high:
            year_low = year_low + 2
        else:
            print("::::::::::::::::::::DONE: Accessed Last year::::::::::::::::::::\n")


def process_data(data, page_number, year_st, year_ed):
    # load the workbook
    wb = load_workbook(filename="data_cids.xlsx", data_only=True)

    # Select the Sheet to work with
    sheet = wb['Sheet1']
    row = sheet.max_row + 1

    try:
        for i in range(0, len(data)):
            curr_data = BeautifulSoup(str(data[i]), 'html.parser')
            data_attributes = curr_data.find('div').attrs
            sheet.cell(row=row, column=1).value = str(year_st)
            sheet.cell(row=row, column=2).value = str(year_ed)
            sheet.cell(row=row, column=3).value = str(page_number)
            try:
                sheet.cell(row=row, column=4).value = data_attributes['data-cid']
            except Exception as e:
                print(e)
                print("Exception! No DATA_CID found on Tab: " + str(page_number) + " Year Range: " + str(year_st) +
                      " - " + str(year_ed) + " Item number: " + str(i))
                wb.save("data_cids.xlsx")
                pass
            paper_title = ""
            try:
                title = curr_data.find_all('h3', attrs={'class': 'gs_rt'})
                for j in range(0, len(title)):
                    list_data = BeautifulSoup(str(title[j]), 'html.parser')
                    paper_title_dom = list_data.find('a', attrs={'data-clk': re.compile(r".*")})
                    paper_title = paper_title_dom.string
                if paper_title != "":
                    sheet.cell(row=row, column=5).value = paper_title
            except:
                try:
                    title = curr_data.find('h3', attrs={'class': 'gs_rt'})
                    sheet.cell(row=row, column=5).value = title.text
                except Exception as s:
                    print(s)
                    print("Exception! No Title found on Tab: " + str(page_number) + " Year Range: " + str(year_st)
                          + " - " + str(year_ed) + " Item number: " + str(i))
                    pass
                wb.save("data_cids.xlsx")
                pass
            row = row + 1
    except Exception as e:
        print(e)
        print("Exception Occurred at: " + str(page_number) + " Year Range: " + str(year_st) + " - " + str(year_ed))
        wb.save("data_cids.xlsx")
        pass

    wb.save("data_cids.xlsx")


def get_request(page_scrape):
    r = http.request(
        'GET',
        page_scrape,
        headers={
            "authority": "scholar.google.ca",
            "method": "GET",
            "path": "/scholar?hl=en&as_sdt=0%2C5&sciodt=0%2C5&cites=10992335886072847329&scipsc="
                    "&as_ylo=2007&as_yhi=2018",
            "scheme": "https",
            "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
            "accept-encoding": "gzip, deflate, br",
            "accept-language": "en-US,en;q=0.9",
            "cache-control": "max-age=0",
            "cookie": "CONSENT=YES+CA.en+20180513-19-0; OGPC=19008862-2:19009353-2:; SID=ywaLe08D1fWoWo4vM_"
                      "MuMUY-BD2KKUbxmk1VLh5x7ZU584e8n7z75ggsP5170Eq4iEJKog.; HSID=Agla5niMAI11ewhpo; SSID=ARe"
                      "4bHC3LJ9J1o_G8; APISID=Tx_tbXnNiNGGwWXB/AsFjYajKMY0X5uz9K; SAPISID=3d0QNj5APzigUvfQ/AbQ"
                      "HLnnRNIUo23otX; 1P_JAR=2018-12-06-14; GSP=LM=1544106761:S=jwEPyXj5FmtWoSnG; NID=150=kdX"
                      "wo88eCR7VONF7hSdQwqbw9Z-W-v439oY7rtFgOdTa2UH2UnnUr690ycdDKoB-8tI6a5H9ApKlIyispXpX4GDdOp"
                      "ECEpfnXP0Hi7SA1aTgqFZCwsbuD3k5Guv598iSm3TEUVPtm5e7-cMf3jIQNQxlad8vltxSjpxwAsnxYOnfDVP_tn"
                      "EiGxw_YCLmd7XfAODr6Jy9clIc28XoBhNR0alh0wTr8nTlpwGKOLDXH2AUj9X-LyslADSBygCqKSFYJfyI87LkdN"
                      "qVW_DyQ50dZosmA7FKQJN2ZPQ",
            "upgrade-insecure-requests": "1",
            "user-agent": "Mozilla/5.0 (Linux; Android 6.0; P027 Build/MRA58L) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/57.0.2987.132 Safari/537.36",
            "x-client-data": "CJK2yQEIprbJAQjEtskBCKmdygEIqKPKAQi/p8oBGPmlygE="
        }
    )

    return r


def process_request(r):
    soup = BeautifulSoup(r.data, 'html.parser')
    citation_papers = soup.find_all('div', attrs={'class': 'gs_r gs_or gs_scl'})
    print(soup)
    return citation_papers


if __name__ == '__main__':
    main()