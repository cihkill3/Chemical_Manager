import sds_downloader
from seleniumbase import Driver
import os

save_dir = 'sds_downloads'
os.makedirs(save_dir, exist_ok=True)

driver = Driver(uc=True, headless=True)

print("Testing 704060F...")
res_tf = sds_downloader.download_thermofisher_sds('704060F', save_dir)
print('TF:', res_tf)
res_al = sds_downloader.download_aldrich_sds('704060F', 'aldrich', save_dir)
print('Aldrich:', res_al)
res_tci = sds_downloader.download_tci_sds('704060F', save_dir)
print('TCI:', res_tci)

print("Testing A10436...")
res_tf = sds_downloader.download_thermofisher_sds('A10436', save_dir)
print('TF:', res_tf)
res_al = sds_downloader.download_aldrich_sds('A10436', 'aldrich', save_dir)
print('Aldrich:', res_al)
res_tci = sds_downloader.download_tci_sds('A10436', save_dir)
print('TCI:', res_tci)

driver.quit()
