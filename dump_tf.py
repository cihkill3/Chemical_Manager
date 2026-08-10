from seleniumbase import Driver
import time

def dump():
    driver = Driver(uc=True, headless=True)
    driver.get("https://www.thermofisher.com/order/catalog/product/L09319")
    time.sleep(10)
    with open("thermo_body.txt", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    driver.quit()

if __name__ == "__main__":
    dump()
