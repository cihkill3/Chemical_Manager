//TOK-1050:
   	 	$(".js-specification").on("click", function(e) {
   			console.log("Specification Link clicked");
   			e.preventDefault();
   			$('.mm_stock_err_msg').each(function() {
               $(this).addClass('hide');
            });
   			$("#acDocNoProductError").hide();
   			$("#sdsAcDocErrorProduct").hide();
   			$("#errorLotNumberAcDoc").hide();
   			$("#sdsCofaErrorProduct").hide();
   			$("#errorLotNumber").hide();
   			$("#cofa-error").hide();
   			$("#errorProductspec").hide();
   			$("#specSearchErrorMsg").hide();
   			$("#sdsNoProductError").hide();
   			$("#specification-error").hide();
   			$("#sample-CofA-error").hide();
   			var url = ACC.config.encodedContextPath + "/documentSearch/productInfoSpecDocFromPDP";
            var productCode = $(".js-specification-productCode").val();
            var brandCode = $("#brandSelector").val().toUpperCase();
   	        
   	        console.log("Brand Code:", brandCode, "Product Code:", productCode, "URL:", url);
   	               
   	               $.post({
   	         		  url :url,
   	         		  data: {brandCode: brandCode, productCode: productCode},
   	         		  xhrFields: {responseType: 'blob'},
   	         	    })
   	         	   .done(function (response, status, xhr) {
   	         		  console.log("Success");
   	         		  var filename = "";                   
   	                  var disposition = xhr.getResponseHeader('Content-Disposition');

   	                  if (disposition) {
   	                     var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
   	                     var matches = filenameRegex.exec(disposition);
   	                     if (matches !== null && matches[1]) filename = matches[1].replace(/['"]/g, '');
   	                  } 
   	                  var linkelem = document.createElement('a');
   	                  try {
   	                    var blob = new Blob([response], { type: 'application/octet-stream' });                        

   	                     if (typeof window.navigator.msSaveBlob !== 'undefined') {
   	                         //   IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
   	                         window.navigator.msSaveBlob(blob, filename);
   	                     } else {
   	                         var URL = window.URL || window.webkitURL;
   	                         var downloadUrl = URL.createObjectURL(blob);

   	                         if (filename) { 
   	                             // use HTML5 a[download] attribute to specify filename
   	                             var a = document.createElement("a");

   	                             // safari doesn't support this yet
   	                             if (typeof a.download === 'undefined') {
   	                                 window.location = downloadUrl;
   	                             } else {
   	                                 a.href = downloadUrl;
   	                                 a.download = filename;
   	                                 document.body.appendChild(a);
   	                                 a.target = "_blank";
   	                                 a.click();
   	                             }
   	                         } else {
   	                             window.location = downloadUrl;
   	                         }
   	                     }   
   	                 } catch (ex) {
   	                     console.log(ex);
   	                   } 
   	         		})
   	         	   .fail(function (jqXHR, textStatus, errorThrown) {
   	         		  console.log("Error details [" + jqXHR.status + ", " + textStatus + ", " + errorThrown + "]");
   	         		  $("#specification-error").show();
   	         		})
   	 	});
   	 	
   	 	$(".js-cofa").on("click", function(e) {
   	 		console.log("C Of A button clicked");
   	 		    var stockErrMsg = $('.js-cofa').siblings('.mm_stock_err_msg').eq(0);
                $('.mm_stock_err_msg').each(function() {
                   $(this).addClass('hide');
                });
   	 		    $("#acDocNoProductError").hide();
   	 		    $("#sdsAcDocErrorProduct").hide();
   	 		    $("#errorLotNumberAcDoc").hide();
   	 		    $("#errorProductspec").hide();
   	 		    $("#sdsErrorProduct").hide();
   	 		    $("#sdsNoProductError").hide();
   	 		    $("#specSearchErrorMsg").hide();
   	 		    $("#specification-error").hide();
   	 		    $("#sample-CofA-error").hide();
   	 		    $("#cofa-error").hide();
		   	 	if ($("#ProductCodeCofa").val() === '') {
		   	 	   $("#sdsCofaErrorProduct").show();
		   	 	   $("#ProductCodeCofa").focus();
		   	 	   return false;
		   	 	} else {
		   	 	   $("#sdsCofaErrorProduct").hide();
		   	 	}


		   	 	if ($("#LotNumbere").val() === '') {
					$("#errorLotNumber").show();
					$("#LotNumbere").focus();
					return false;
				} else {
					$("#errorLotNumber").hide();
				}
		
				re = /^[A-Za-z0-9]+$/;
				if (!re.test($("#LotNumbere").val())) {
					$("#errorLotNumberAlphaNum").show();
					$("#LotNumbere").focus();
					return false;
				} else {
					$("#errorLotNumberAlphaNum").hide();
				}
            $(".js-cofa").addClass('hide');
            $(this).siblings('.btn-loading').removeClass('hide');
            $('.lot-suggestion-list').addClass('hide');
            $('.lot-suggestion-list .option').remove();

			var url;
			var certificateType = document.getElementById("Type").value;
				if (certificateType == 'COA') {
					url = ACC.config.encodedContextPath + "/documentSearch/productInfoCofAFromPDP";
			} else if (certificateType == 'COO') {
				url = ACC.config.encodedContextPath + "/documentSearch/productInfoCOO";
			} else if (certificateType == 'BSETSE') {
				url = ACC.config.encodedContextPath + "/documentSearch/productInfoBSETSE";
			} else if (certificateType == 'CCOO') {
				url = ACC.config.encodedContextPath + "/documentSearch/productInfoCCOO";
			}
			console.log("URL "+url); 
   			var productCode = $(".js-productCode-Cofa").val();
   	        var lotNumber = $(".js-lotNumber").val();
            var brandCode = $("#brandSelector").val().toUpperCase();
   	        console.log("Product Code:", productCode, "Lot Number:",lotNumber);

   	        	$.post({
   		    		  url :url,
   		    		  data: {brandCode: brandCode, productCode: productCode, lotNumber: lotNumber},
   		    		  xhrFields: {responseType: 'blob'}
   		    	    })
   		    	   .done(function (response, status, xhr) {
   		    		   $("#cofa-error").hide();
   		    		   stockErrMsg.addClass("hide");
   		    		   console.log("Success");
   	         		   var filename = "";                   
   	                   var disposition = xhr.getResponseHeader('Content-Disposition');

   	                  if (disposition) {
   	                     var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
   	                     var matches = filenameRegex.exec(disposition);
   	                     if (matches !== null && matches[1]) filename = matches[1].replace(/['"]/g, '');
   	                  } 
   	                  var linkelem = document.createElement('a');
   	                  try {
   	                    var blob = new Blob([response], { type: 'application/octet-stream' });                        

   	                     if (typeof window.navigator.msSaveBlob !== 'undefined') {
   	                         //   IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
   	                         window.navigator.msSaveBlob(blob, filename);
   	                     } else {
   	                         var URL = window.URL || window.webkitURL;
   	                         var downloadUrl = URL.createObjectURL(blob);

   	                         if (filename) { 
   	                             // use HTML5 a[download] attribute to specify filename
   	                             var a = document.createElement("a");

   	                             // safari doesn't support this yet
   	                             if (typeof a.download === 'undefined') {
   	                                 window.location = downloadUrl;
   	                             } else {
   	                                 a.href = downloadUrl;
   	                                 a.download = filename;
   	                                 document.body.appendChild(a);
   	                                 a.target = "_blank";
   	                                 a.click();
   	                             }
   	                         } else {
   	                             window.location = downloadUrl;
   	                         }
   	                     }   
   	                 } catch (ex) {
   	                     console.log(ex);
   	                   } 
   	         		})
   		    	   .fail(function (response, status, xhr) {
   		    	        console.log("failed");
                        let responseText = getResponseText(response);
                        let suggestionList = $(e.target).closest('.js-cofa_pdp').find('.lot-suggestion-list');
                        if(responseText != null && responseText == 'MM_OUT_OF_STOCK'){
                            console.log("MM_OUT_OF_STOCK");
                            stockErrMsg.removeClass("hide");
                            $("#cofa-error").hide();
                        }
                        else if(responseText != null && suggestionList) {
                            responseText = responseText.slice(1, -1);
                            let arr = responseText.split(',').map(item => item.trim());
                            arr.forEach(item => {
                              suggestionList.append(`<div class="option" data-value="${item}">${item}</div>`);
                            });
                            suggestionList.removeClass('hide');
                            $('#LotNumbere').focus();
                        }
                        else {
   		    	            var lotNumberValue = $(".js-lotNumber").val();
                            var productCodeValue = $(".js-productCode-Cofa").val();
                            var link = ACC.config.encodedContextPath + "/contact-us"
                            $("#lotNumberLink").attr("href", link+''+'?lotNumber='+lotNumberValue+'&productCode='+productCodeValue);
                            $("#cofa-error").show();
                            stockErrMsg.addClass("hide");
   		    	        }
   		    	    })
   		    		.always(function() {
                        $(".js-cofa").removeClass('hide');
                        $(".js-cofa").siblings('.btn-loading').addClass('hide');
                    });
   	 		});
   	 		
   	 //TOK-1050


		//TICKET-31437
		//getProductAnalyticChartSearchDocsForALlType
		const fileTypeMap = {
			GC: 'GC',
			LC: 'LC',
			H: '¹H NMR',
			C: '¹³C NMR',
			F: '¹⁹F NMR',
			P: '³¹P NMR',
			IR: 'IR'
		};
		$('.js-AcDoc').on("click", function(e) {
			e.preventDefault();
			const $btn = $(this);
			const stockErrMsg = $btn.siblings('.mm_stock_err_msg').eq(0);

			$('.mm_stock_err_msg').addClass('hide');
			$("#sdsCofaErrorProduct, #errorLotNumber, #cofa-error, #errorProductspec, #sdsErrorProduct, #sdsNoProductError, #specSearchErrorMsg, #specification-error, #sample-CofA-error, #acDocNoProductError").hide();

			if ($("#ProductCodeAcDoc").val() === '') {
				$("#sdsAcDocErrorProduct").show();
				$("#ProductCodeAcDoc").focus();
				return false;
			} else {
				$("#sdsAcDocErrorProduct").hide();
			}

			if ($("#LotNumberAcDoc").val() === '') {
				$("#errorLotNumberAcDoc").show();
				$("#LotNumberAcDoc").focus();
				return false;
			} else {
				$("#errorLotNumberAcDoc").hide();
			}

			let re = /^[A-Za-z0-9]+$/;
			if (!re.test($("#LotNumberAcDoc").val())) {
				$("#errorLotNumberAlphaNumAcDoc").show();
				$("#LotNumbereAcDoc").focus();
				return false;
			} else {
				$("#errorLotNumberAlphaNumAcDoc").hide();
			}

			$(".js-AcDoc").addClass('hide');
			$(this).siblings('.btn-loading').removeClass('hide');
			$('.lot-suggestion-list').addClass('hide').find('.option').remove();

			const url = ACC.config.encodedContextPath + "/documentSearch/productAcDocAllTypes";
			const productCode = $(".js-productCode-AcDoc").val();
			const lotNumber = $(".js-lotNumber-AcDoc").val();
            const brandCode = $("#brandSelector").val().toUpperCase();

			$.post({
				url: url,
				data: {
                    brandCode: brandCode,
					productCode: productCode,
					lotNumber: lotNumber
				}
			}).then(function (response, status, xhr) {
				const msg = response.responseJSON?.msg || response.msg;
				let suggestionList = $(e.target).closest('.js-AcDoc_pdp').find('.lot-suggestion-list');
				// If msg is a specific error string
				if (msg === 'success') {
					$("#sdsErrorProduct, #sdsNoProductError").hide();
					stockErrMsg.addClass("hide");

					if (response.productDetailsList && response.productDetailsList.length > 0) {
						// Set product and lot info
						$('#popup-product-code').text(response.productNo);
						$('#popup-lot-number').text(response.lotNo);

						// Clear old links
						$('#popup-file-links').empty();
						const acurl = ACC.config.encodedContextPath + "/documentSearch/productAcDoc";

						// Add download links using dynamic URLs
						response.productDetailsList.forEach((file, index) => {
							const fileType = encodeURIComponent(file || `File${index + 1}`);
							const $link = $('<a>')
								.addClass('arial-link')
								.attr('href', acurl + "?brandCode=" + brandCode + "&productCode=" + encodeURIComponent(productCode) + "&lotNumber=" + encodeURIComponent(lotNumber) + "&type=" + fileType + "&download=false")
								.text(fileTypeMap[file] || `File ${index + 1}`);

							const $li = $('<li>').append($link);
							$('#popup-file-links').append($li);
						});

						// Show popup
						$('#acdoc-popup').modal('show');
						$('#acdoc-popup').removeClass('cboxElement');
					}
				}
				else if (msg === 'MM_OUT_OF_STOCK') {
					console.log('MM_OUT_OF_STOCK');
					stockErrMsg.removeClass("hide");
					$("#acDocNoProductError").hide();
				}
				// If msg is a comma-separated suggestion list
				else if (typeof msg === 'string' && msg.includes('[') && suggestionList) {
					let sugLots = msg.slice(1, -1);
					const arr = sugLots.split(',').map(item => item.trim());
					const suggestionList = $btn.closest('.js-AcDoc_pdp').find('.lot-suggestion-list');
					arr.forEach(item => {
						suggestionList.append(`<div class="option" data-value="${item}">${item}</div>`);
					});
					suggestionList.removeClass('hide');
					$('#LotNumberAcDoc').focus();
				}
				// If msg is an object with data to show in modal

				// If msg is unknown
				else {
					$("#acDocNoProductError").show();
					stockErrMsg.addClass("hide");
				}
			}).catch(function (err) {
					console.error("Unexpected error", err);
					$("#acDocNoProductError").show();
			}).always(function() {
				$(".js-AcDoc").removeClass('hide');
				$(".js-AcDoc").siblings('.btn-loading').addClass('hide');
			});
		});


// SDS Search
$("#sdsSearchButton").on("click", function (e) {
    e.preventDefault();

    var isKISHIDAProduct = ($("#brandSelector").val() || "").toUpperCase() === "KISHIDA";

    $(".mm_stock_err_msg").each(function () {
        $(this).addClass("hide");
    });

    // Other forms
    $("#acDocNoProductError").hide();
    $("#sdsAcDocErrorProduct").hide();
    $("#errorLotNumberAcDoc").hide();
    $("#sdsCofaErrorProduct").hide();
    $("#errorLotNumber").hide();
    $("#cofa-error").hide();
    $("#errorProductspec").hide();
    $("#specSearchErrorMsg").hide();
    $("#specification-error").hide();
    $("#sample-CofA-error").hide();

    // Form validation
    if ($("#sdsProductCode").val() === '') {
        $("#sdsNoProductError").hide();
        $("#sdsErrorProduct").show();
        $("#sdsErrorPartnerProductInvalid").hide();
        $("#langSelectorError").hide();
        $("#sdsProductCode").focus();
        return false;
    } else if ($("#langSelector").val() === '') {
        $("#sdsNoProductError").hide();
        $("#sdsErrorProduct").hide();
        $("#sdsErrorPartnerProductInvalid").hide();
        $("#langSelectorError").show();
        $("#langSelector").focus();
        return false;
    } else if (isKISHIDAProduct && ($("#sdsProductCode").val().length !== 8 && $("#sdsProductCode").val().length !== 9)) { // TICKET-36666 - Only check product KISHIDA when brandSelector is KISHIDA and productCode length is 8 or 9
        $("#sdsNoProductError").hide();
        $("#sdsErrorProduct").hide();
        $("#sdsErrorPartnerProductInvalid").show();
        $("#langSelectorError").hide();
        $("#sdsProductCode").focus();
        return false;
    } else {
        $("#sdsNoProductError").hide();
        $("#sdsErrorProduct").hide();
        $("#sdsErrorPartnerProductInvalid").hide();
        $("#langSelectorError").hide();
    }

    // ajax call to productSDSSearchDoc.
    var url = ACC.config.encodedContextPath + "/documentSearch/productSDSSearchDoc";
    var brandCode = $("#brandSelector").val().toUpperCase();
    var productCode = $("#sdsProductCode").val().toUpperCase();
    var langSelector = $("#langSelector").val();
    var selectedCountry = $("#selectedCountry").val();
    console.log(productCode + " " + langSelector + " " + selectedCountry);
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
        console.log("Done");
        $("#sdsNoProductError").hide();
        $("#sdsErrorProduct").hide();
        $("#sdsErrorPartnerProductInvalid").hide();
        $("#langSelectorError").hide();
        var filename = "";
        var disposition = xhr.getResponseHeader('Content-Disposition');
        console.log(disposition);
        if (disposition) {
            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
            var matches = filenameRegex.exec(disposition);
            if (matches !== null && matches[1]) {
                filename = matches[1].replace(/['"]/g, '');
                const blob = new Blob([response], {type: 'application/pdf'});
                const downloadUrl = window.URL.createObjectURL(blob);
                console.log(downloadUrl);
                if (filename) {
                    var a = document.createElement("a");
                    // safari doesn't support this yet
                    if (typeof a.download === 'undefined') {
                        window.location = downloadUrl;
                    } else {
                        a.href = downloadUrl;
                        a.download = filename;
                        document.body.appendChild(a);
                        a.target = "_blank";
                        a.click();
                        document.body.removeChild(a);
                    }
                } else {
                    window.location = downloadUrl;
                }
                window.URL.revokeObjectURL(downloadUrl);
            }
        }
        console.log("File Downloaded");
    }).fail(function (response, status, xhr) {
        console.log(url);
        console.log("Fail");
        var productCode = $("#sdsProductCode").val();
        var langSelector = $("#langSelector option:selected").text().trim();
        var link = $("#contactUsLinkValue").val();
        $("#sdsProductCodeLink").attr("href", link + '?' + 'language=' + langSelector + '&productCode=' + productCode);
        $("#sdsNoProductError").show();
    })
});

function getResponseText(response){
   if(response != null && response.getAllResponseHeaders() != null){
        let headers = response.getAllResponseHeaders();
        let responseTextValue = headers.split('\r\n').find(header => header.startsWith('response-text:'))?.split(': ')[1];
        return responseTextValue;
   }
   return null;
}

$('.lot-suggestion-list').on('mousedown', '.option', function(e) {
    e.preventDefault();
    let coaParents = $(e.target).closest('.js-cofa_pdp');
    let acParents = $(e.target).closest('.js-AcDoc_pdp');
    let lotSelected = $(this).data('value');
    if(coaParents){
        coaParents.find('#LotNumbere').val(lotSelected);
        coaParents.find('.js-cofa').trigger('click');
    }
    if(acParents){
        acParents.find('#LotNumberAcDoc').val(lotSelected);
        acParents.find('.js-AcDoc').trigger('click');
    }
});

$('#LotNumbere, #LotNumberAcDoc').on('input', function(e){
    if(!$(this).val()){
        $('.lot-suggestion-list').addClass('hide');
    }
});

$('#Type, #AcDocType').on('change', function(e){
    if($(this).val()){
        $('.lot-suggestion-list').addClass('hide');
        $('.lot-suggestion-list .option').remove();
    }
});
$('#LotNumbere, #LotNumberAcDoc').on('focusin', function (e) {
    let coaParents = $(e.target).closest('.js-cofa_pdp');
    let acParents = $(e.target).closest('.js-AcDoc_pdp');
    if($(this).val() && coaParents.find('.lot-suggestion-list').find('.option').length > 0){
        coaParents.find('.lot-suggestion-list').removeClass('hide');
    }
    if($(this).val() && acParents.find('.lot-suggestion-list').find('.option').length > 0){
        acParents.find('.lot-suggestion-list').removeClass('hide');
    }
});
$('#LotNumbere, #LotNumberAcDoc').on('focusout', function (e) {
    $('.lot-suggestion-list').addClass('hide');
});

/**
TICKET-29881: [Global] Add sample COA download function on PDP.
**/
$(".js-sample-cofa").on("click", function(e) {
	console.log("Sample C Of A button clicked");

	$('.mm_stock_err_msg').each(function() {
	   $(this).addClass('hide'); 				// hide mmstock error message
	});
	$("#acDocNoProductError").hide();			// hide AC not found error message
	$("#errorLotNumberAcDoc").hide();			// hide AC no product lot input error message
	$("#errorLotNumberAlphaNumAcDoc").hide();	// hide AC invalid product lot input error message
	$("#sdsNoProductError").hide();				// hide SDS not found error message
	$("#specSearchErrorMsg").hide();			// hide Spec not found error message
	$("#cofa-error").hide();					// hide CoA not found error message
	$("#errorLotNumber").hide();				// hide CoA no product lot input error message
	$("#errorLotNumberAlphaNum").hide();		// hide CoA invalid product lot input error message

	$.post({
	  url : ACC.config.encodedContextPath + "/documentSearch/sampleCofAFromPDP",
	  data: {
          brandCode: $("#brandSelector").val(),
          productCode: $(".js-productCode-Cofa").val()
      },
	  xhrFields: {responseType: 'blob'}
	})
	.done(function (response, status, xhr) {
		$("#sample-CofA-error").hide();
		console.log("Success");
		var filename = "";
		var disposition = xhr.getResponseHeader('Content-Disposition');

		if (disposition) {
			var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
			var matches = filenameRegex.exec(disposition);
			if (matches !== null && matches[1]) filename = matches[1].replace(/['"]/g, '');
		}
		try {
			var blob = new Blob([response], { type: 'application/octet-stream' });
			if (typeof window.navigator.msSaveBlob !== 'undefined') {
				 window.navigator.msSaveBlob(blob, filename);
			} else {
				 var URL = window.URL || window.webkitURL;
				 var downloadUrl = URL.createObjectURL(blob);
				 var a = document.createElement("a");
				 if (filename && typeof a.download !== 'undefined') {
                     a.href = downloadUrl;
                     a.download = filename;
                     document.body.appendChild(a);
                     a.target = "_blank";
                     a.click();
				 } else {
					 window.location = downloadUrl;
				 }
			}
		} catch (ex) {
			console.log(ex);
			$("#sample-CofA-error").show();
		}
	})
	.fail(function (response, status, xhr) {
		console.log("failed");
		$("#sample-CofA-error").show();
	})
});

function applyBrandSelectorLogic() {
    const brand = $("#brandSelector").val();

    const productCodeMaxLength = (brand === "KISHIDA") ? 9 : 5;
    const lotNumberMaxLength   = (brand === "KISHIDA") ? 7 : 5;

    // Product code related inputs
    const $productCodeInputs = $(
        "#sdsProductCode, " +
        "#text\\.certificate\\.analysis\\.productNumber, " +
        "#ProductCodeCofa, " +
        "#ProductCodeAcDoc"
    );

    $productCodeInputs.each(function () {
        const $input = $(this);

        // ----------------------------
        // 1. Set maxlength
        // ----------------------------
        $input.attr("maxlength", productCodeMaxLength);

        const val = $input.val();
        if (val && val.length > productCodeMaxLength) {
            $input.val(val.substring(0, productCodeMaxLength));
        }

        // ----------------------------
        // 2. Update placeholder by brand (from hidden inputs)
        // ----------------------------
        const placeholderKey = brand === "KISHIDA"
            ? "#input\\.productNum\\.placeholder\\.KISHIDA"
            : "#input\\.productNum\\.placeholder\\.TCI";

        const brandPlaceholder = $(placeholderKey).html();

        if (brandPlaceholder) {
            $input.attr("placeholder", brandPlaceholder.trim());
        }
    });

    // Lot number inputs
    const $lotNumberInputs = $("#LotNumbere, #LotNumberAcDoc");

    $lotNumberInputs.each(function () {
        const $input = $(this);

        // ----------------------------
        // 1. Set maxlength
        // ----------------------------
        $input.attr("maxlength", lotNumberMaxLength);

        const val = $input.val();
        if (val && val.length > lotNumberMaxLength) {
            $input.val(val.substring(0, lotNumberMaxLength));
        }

        // ----------------------------
        // 2. Update lot number placeholder (from hidden inputs)
        // ----------------------------
        const lotPlaceholderKey = brand === "KISHIDA"
            ? "#input\\.lotNum\\.placeholder\\.KISHIDA"
            : "#input\\.lotNum\\.placeholder\\.TCI";

        const lotPlaceholder = $(lotPlaceholderKey).html();

        if (lotPlaceholder) {
            $input.attr("placeholder", lotPlaceholder.trim());
        }
    });

    // Other elements (error messages, labels, etc.)
    const brandSuffix = brand === "KISHIDA" ? "KISHIDA" : "TCI";

    /* -----------------------------
     * sdsNoProductError
     * ----------------------------- */
    const $sdsNoProductError = $("#sdsNoProductError");
    const $sdsNoProductErrorBrand = $("#sdsNoProductError\\." + brandSuffix);

    if ($sdsNoProductErrorBrand.length) {
        $sdsNoProductError.html($sdsNoProductErrorBrand.html().trim());
    }

    /* -----------------------------
     * specSearchErrorMsg
     * ----------------------------- */
    const $specSearchErrorMsg = $("#specSearchErrorMsg");
    const $specSearchErrorMsgBrand = $("#specSearchErrorMsg\\." + brandSuffix);

    if ($specSearchErrorMsgBrand.length) {
        $specSearchErrorMsg.html($specSearchErrorMsgBrand.html().trim());
    }

    /* -----------------------------
     * specification-error
     * ----------------------------- */
    const $specificationError = $("#specification-error");
    const $specificationErrorBrand = $("#specification-error\\." + brandSuffix);

    if ($specificationErrorBrand.length) {
        $specificationError.html($specificationErrorBrand.html().trim());
    }

    /* -----------------------------
     * errorLotNumberAlphaNum
     * ----------------------------- */
    const $errorLotNumber = $("#errorLotNumberAlphaNum");
    const $errorLotNumberBrand = $("#errorLotNumberAlphaNum\\." + brandSuffix);

    if ($errorLotNumberBrand.length) {
        $errorLotNumber.html($errorLotNumberBrand.html().trim());
    }

    /* -----------------------------
     * errorLotNumberAlphaNumAcDoc
     * ----------------------------- */
    const $errorLotNumberAcDoc = $("#errorLotNumberAlphaNumAcDoc");
    const $errorLotNumberBrandAcDoc = $("#errorLotNumberAlphaNumAcDoc\\." + brandSuffix);

    if ($errorLotNumberBrandAcDoc.length) {
        $errorLotNumberAcDoc.html($errorLotNumberBrandAcDoc.html().trim());
    }

    /* -----------------------------
     * cofa-error
     * ----------------------------- */
    const $cofaError = $("#cofa-error");
    const $cofaErrorBrand = $("#cofa-error\\." + brandSuffix);

    if ($cofaErrorBrand.length) {
        $cofaError.html($cofaErrorBrand.html().trim());
    }

    // Type select logic
    const $typeSelect = $("#Type");

    if (brand === "KISHIDA") {
        $typeSelect.val("COA");

        $typeSelect.find("option").each(function () {
            if ($(this).val() !== "COA") {
                $(this).prop("disabled", true).hide();
            }
        });

        $typeSelect.prop("disabled", true);
    } else {
        $typeSelect.prop("disabled", false);
        $typeSelect.find("option").prop("disabled", false).show();
    }
}

$(document).ready(function () {
    $("#brandSelector").on("change", applyBrandSelectorLogic);

    // init render
    const params = new URLSearchParams(window.location.search);
    const brandCode = params.get("brandCode");

    if (brandCode) {
        $("#brandSelector").val(brandCode);
    }

    applyBrandSelectorLogic();
});