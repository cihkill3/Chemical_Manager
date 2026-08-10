"""
TCI SDS Direct jQuery $.post test inside browser
"""
import os
import time
import base64
from seleniumbase import Driver

def test_tci_direct_post(product_number: str):
    d = Driver(uc=True, headless=True)
    url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
    print(f"[{product_number}] 접속 중...")
    d.get(url)
    time.sleep(3)
    
    # 브라우저 내부에서 jQuery $.post 직접 호출 테스트
    script = """
    return new Promise((resolve, reject) => {
        var url = ACC.config.encodedContextPath + "/documentSearch/productSDSSearchDoc";
        var brandCode = $("#brandSelector").val() ? $("#brandSelector").val().toUpperCase() : "TCI";
        var productCode = $("#sdsProductCode").val() ? $("#sdsProductCode").val().toUpperCase() : "%s";
        var langSelector = $("#langSelector").val() || "EN";
        var selectedCountry = $("#selectedCountry").val() || "KR";
        
        console.log("POSTing to:", url, {brandCode, productCode, langSelector, selectedCountry});
        
        $.post({
            url: url,
            data: {
                brandCode: brandCode,
                productCode: productCode,
                langSelector: langSelector,
                selectedCountry: selectedCountry
            },
            xhrFields: {responseType: 'blob'}
        }).done(function (response, status, xhr) {
            var reader = new FileReader();
            reader.readAsDataURL(response);
            reader.onloadend = function() {
                resolve({
                    status: 'success',
                    dataUrl: reader.result,
                    disposition: xhr.getResponseHeader('Content-Disposition')
                });
            };
        }).fail(function (jqXHR, textStatus, errorThrown) {
            resolve({
                status: 'fail',
                httpStatus: jqXHR.status,
                textStatus: textStatus,
                errorThrown: errorThrown,
                responseText: jqXHR.responseText
            });
        });
    });
    """ % product_number

    res = d.execute_script(script)
    print("Direct $.post 결과:", res.get("status"))
    if res.get("status") == "success":
        data_url = res.get("dataUrl", "")
        print(f"Content-Disposition: {res.get('disposition')}")
        print(f"Data URL length: {len(data_url)}")
        if "," in data_url:
            b64 = data_url.split(",")[1]
            pdf_bytes = base64.b64decode(b64)
            print(f"PDF Size: {len(pdf_bytes)} bytes")
            with open(f"SDS_TCI_{product_number}.pdf", "wb") as f:
                f.write(pdf_bytes)
            print(f"저장 성공: SDS_TCI_{product_number}.pdf")
    else:
        print("Fail detail:", res)
        
    d.quit()

if __name__ == "__main__":
    test_tci_direct_post("C0119")
