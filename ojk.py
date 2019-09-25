import os
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

os.system('clear')

def every_downloads_chrome(driver):
    if not driver.current_url.startswith("chrome://downloads"):
        driver.get("chrome://downloads/")
    return driver.execute_script("""
        var items = downloads.Manager.get().items_;
        if (items.every(e => e.state === "COMPLETE"))
            return items.map(e => e.fileUrl || e.file_url);
        """)

with open('bulan.json') as json_bulan:
    bulan = json.load(json_bulan)
    data_bulan = bulan["data"]
    kode_bulan = []
    print("[+] Pilih Bulan Laporan : \n")
    for b in data_bulan:
        kode_bulan.append(b["MonthCode"])
        print("[{}]\t {}".format(b["MonthCode"],b["MonthName"]))

    pilihan_bulan = input("\n[+] Bulan Laporan : ")
    pilihan_tahun = input("[+] Tahun Laporan : ")

    if pilihan_bulan not in kode_bulan :
        print("[!] Bulan yang tersedia hanya 3,6,9 dan 12\n")
        os.system("exit")
    else:
        with open('laporan.json') as json_laporan:
            laporan = json.load(json_laporan)
            data_laporan = laporan["data"]
            print("\n[+] Pilih Jenis Laporan :\n")
            for lap in data_laporan:
                print("[{}] {} {}".format(lap["kode"], lap["id"], lap["text"]))
        pilihan_laporan = input("\n[+] Laporan : ")

        laporan_terpilih = [items for items in data_laporan if items["kode"] == pilihan_laporan ]
        bulan_terpilih  = [items for items in data_bulan if items["MonthCode"] == pilihan_bulan ]

        print("\n[!] Aplikasi Akan Mendownload {} untuk bulan {} tahun {}".format(laporan_terpilih[0]["text"],bulan_terpilih[0]["MonthName"],pilihan_tahun))
        with open('bank_mini.json') as daftar_bank:
            daftar = json.load(daftar_bank)
            print("[!] Ditemukan {} Daftar Bank".format(len(daftar["data"])))
            yakin = input("[!] Proses download akan membutuhkan waktu, apakah anda yakin [y/n] ? : ")
            if yakin == "y":
                print("\n[+] Sedang Mendownload....\n")
                for bank in daftar["data"]:
                    nmBank = bank["text"].replace(" ","+")
                    folder = "{}/{}/{}/{}".format(laporan_terpilih[0]["text"],pilihan_tahun,bulan_terpilih[0]["MonthName"],bank["text"])
                    if not os.path.exists(folder):
                        os.makedirs(folder)
                    uri = "https://cfs.ojk.go.id/cfs/ReportViewerForm.aspx?BankCode={}&Month={}&Year={}&FinancialReportTypeCode={}".format(nmBank,bulan_terpilih[0]["MonthCode"],pilihan_tahun,laporan_terpilih[0]["id"])
                    option = webdriver.ChromeOptions()
                    option.add_argument("--incognito")
                    #option.add_argument('--headless')
                    #option.add_argument('--disable-gpu')
                    option.add_argument("--window-size=400,400")
                    option.add_argument('disable-component-cloud-policy')
                    option.add_experimental_option("prefs", {
                      "download.default_directory": r"{}".format(os.getcwd()+"/"+folder),
                      "download.prompt_for_download": False,
                      "download.directory_upgrade": False,
                      "safebrowsing.enabled": True
                    })
                    browser = webdriver.Chrome(executable_path='/Users/khairuaqsarasudirman/Desktop/OJK/driver/chromedriver', options=option)
                    browser.get(uri)
                    timeout = 20
                    try:
                        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="CFSReportViewer_ctl05_ctl04_ctl00_ButtonImg"]')))
                    except TimeoutException:
                        print("[!] Data {} tidak tersedia, atau koneksi internet terputus".format(bank["text"]))
                        browser.quit()

                    path = os.getcwd()+"/"+folder
                    time.sleep(2)
                    try:
                        browser.execute_script("$find('CFSReportViewer').exportReport('EXCELOPENXML');")
                    except Exception as e:
                        print(e)
                    WebDriverWait(browser, 120,1).until(every_downloads_chrome)
                    print("[!] Data {} untuk {} Berhasil Terdownload".format(laporan_terpilih[0]["text"],bank["text"]))
                    browser.quit()
            else:
                os.system("exit")
