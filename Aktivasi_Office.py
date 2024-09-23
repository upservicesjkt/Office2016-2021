import os
import ctypes

# Menentukan judul jendela
ctypes.windll.user32.SetWindowTextW(ctypes.windll.user32.GetForegroundWindow(), "Aktivasi Office 2016-2021 versi AE1 or AE2")

def activate_office(version):
    # Memeriksa apakah folder Licenses16 ada
    if os.path.exists(r"..\root\Licenses16"):
        print("Folder Licenses16 ditemukan, menjalankan konversi Retail ke Volume...")
        os.system("cscript ospp.vbs /dstatus")

        if version == "1":
            print("Menjalankan aktivasi untuk Office 2016...")
            # Tambahkan lisensi untuk Office 2016
            licenses = [
                "ProPlusVL_KMS_Client-ppd.xrm-ms",
                "ProPlusVL_KMS_Client-ul.xrm-ms",
                "ProPlusVL_KMS_Client-ul-oob.xrm-ms",
                "ProPlusVL_MAK-pl.xrm-ms",
                "ProPlusVL_MAK-ppd.xrm-ms",
                "ProPlusVL_MAK-ul-oob.xrm-ms",
                "ProPlusVL_MAK-ul-phn.xrm-ms"
            ]
        elif version == "2":
            print("Menjalankan aktivasi untuk Office 2019...")
            licenses = [
                "ProPlus2019VL_KMS_Client_AE-ppd.xrm-ms",
                "ProPlus2019VL_KMS_Client_AE-ul.xrm-ms",
                "ProPlus2019VL_KMS_Client_AE-ul-oob.xrm-ms",
                "ProPlus2019VL_MAK_AE-pl.xrm-ms",
                "ProPlus2019VL_MAK_AE-ppd.xrm-ms",
                "ProPlus2019VL_MAK_AE-ul-oob.xrm-ms",
                "ProPlus2019VL_MAK_AE-ul-phn.xrm-ms"
            ]
        elif version == "3":
            print("Menjalankan aktivasi untuk Office 2021 AE1...")
            licenses = [
                "ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms",
                "ProPlus2021VL_KMS_Client_AE-ul.xrm-ms",
                "ProPlus2021VL_KMS_Client_AE-ul-oob.xrm-ms",
                "ProPlus2021VL_MAK_AE1-pl.xrm-ms",
                "ProPlus2021VL_MAK_AE1-ppd.xrm-ms",
                "ProPlus2021VL_MAK_AE1-ul-oob.xrm-ms",
                "ProPlus2021VL_MAK_AE1-ul-phn.xrm-ms"
            ]
        elif version == "4":
            print("Menjalankan aktivasi untuk Office 2021 AE2...")
            licenses = [
                "ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms",
                "ProPlus2021VL_KMS_Client_AE-ul.xrm-ms",
                "ProPlus2021VL_KMS_Client_AE-ul-oob.xrm-ms",
                "ProPlus2021VL_MAK_AE2-pl.xrm-ms",
                "ProPlus2021VL_MAK_AE2-ppd.xrm-ms",
                "ProPlus2021VL_MAK_AE2-ul-oob.xrm-ms",
                "ProPlus2021VL_MAK_AE2-ul-phn.xrm-ms"
            ]
        else:
            print("Pilihan tidak valid!")
            return

        for license_file in licenses:
            os.system(f"cscript ospp.vbs /inslic:\"..\\root\\Licenses16\\{license_file}\"")

        # Menggunakan key default untuk konversi awal ke Volume
        default_keys = {
            "1": "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99",
            "2": "NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP",
            "3": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH",
            "4": "FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
        }

        os.system(f"cscript ospp.vbs /inpkey:{default_keys[version]}")

        # Memasukkan product key Volume dari user
        product_key = input("Masukkan product key Volume: ")
        os.system(f"cscript ospp.vbs /inpkey:{product_key}")

        # Mengatur KMS host dan melakukan aktivasi
        os.system("cscript ospp.vbs /sethst:kmstest.contoso.com")
        os.system("cscript ospp.vbs /act")
    else:
        print("Folder Licenses16 tidak ditemukan, menjalankan aktivasi langsung...")
        product_key = input("Masukkan product key Volume: ")
        os.system(f"cscript ospp.vbs /inpkey:{product_key}")
        os.system("cscript ospp.vbs /sethst:kmstest.contoso.com")
        os.system("cscript ospp.vbs /act")

if __name__ == "__main__":
    while True:
        print("====================================")
        print("Pilih versi Office yang ingin diaktivasi:")
        print("Untuk Office 2021 Pastikan Licensi yang dipilih adalah sesuai AE1 atau AE2 :")
        print("1. Office 2016")
        print("2. Office 2019")
        print("3. Office 2021 AE1")
        print("4. Office 2021 AE2")
        print("====================================")
        
        versi = input("Masukkan pilihan [1/2/3/4]: ")

        if versi in ["1", "2", "3", "4"]:
            activate_office(versi)
            break
        else:
            print("Pilihan tidak valid, silakan coba lagi.")

    print("Aktivasi selesai!")
