const W9_COI_FILE_ID = "1U_LvdcogJy1-Ts7AkQytItPelTef4-Nw";

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < 2) return;

  if (sheetName === "Leadlist") {
    const leadIdCol = 1;
    const dateAddedCol = 2;
    const businessNameCol = 3;
    const statusCol = 14;
    const quoteNeededCol = 18;
    const quoteSentCol = 19;
    const respondedCol = 21;
    const scheduledCol = 23;
    const jobCompletedCol = 26;
    const lostLeadCol = 27;
    const priorityCol = 32;

    if (col !== businessNameCol) return;

    const businessName = sheet.getRange(row, businessNameCol).getValue();
    if (!businessName) return;

    const leadIdCell = sheet.getRange(row, leadIdCol);
    if (!leadIdCell.getValue()) {
      leadIdCell.setValue("L-" + new Date().getTime());
    }

    const dateAddedCell = sheet.getRange(row, dateAddedCol);
    if (!dateAddedCell.getValue()) {
      dateAddedCell.setValue(new Date());
    }

    const statusCell = sheet.getRange(row, statusCol);
    if (!statusCell.getValue()) {
      statusCell.setValue("New Lead");
    }

    const quoteNeededCell = sheet.getRange(row, quoteNeededCol);
    if (quoteNeededCell.getValue() === "") {
      quoteNeededCell.setValue(false);
    }

    const quoteSentCell = sheet.getRange(row, quoteSentCol);
    if (quoteSentCell.getValue() === "") {
      quoteSentCell.setValue(false);
    }

    const respondedCell = sheet.getRange(row, respondedCol);
    if (respondedCell.getValue() === "") {
      respondedCell.setValue(false);
    }

    const scheduledCell = sheet.getRange(row, scheduledCol);
    if (scheduledCell.getValue() === "") {
      scheduledCell.setValue(false);
    }

    const jobCompletedCell = sheet.getRange(row, jobCompletedCol);
    if (jobCompletedCell.getValue() === "") {
      jobCompletedCell.setValue(false);
    }

    const lostLeadCell = sheet.getRange(row, lostLeadCol);
    if (lostLeadCell.getValue() === "") {
      lostLeadCell.setValue(false);
    }

    const priorityCell = sheet.getRange(row, priorityCol);
    if (!priorityCell.getValue()) {
      priorityCell.setValue("Medium");
    }

    return;
  }

  if (sheetName === "QuoteLineItems") {
    const lineItemIdCol = 1;
    const quoteIdCol = 2;
    const includedCol = 8;

    if (col !== quoteIdCol) return;

    const quoteId = sheet.getRange(row, quoteIdCol).getValue();
    if (!quoteId) return;

    const lineItemIdCell = sheet.getRange(row, lineItemIdCol);
    if (!lineItemIdCell.getValue()) {
      lineItemIdCell.setValue("LI-" + new Date().getTime());
    }

    const includedCell = sheet.getRange(row, includedCol);
    if (includedCell.getValue() === "") {
      includedCell.setValue(true);
    }

    return;
  }

  const config = {
    "Quotes": {
      idCol: 1,
      triggerCol: 2,
      prefix: "Q-"
    },
    "ServicePricing": {
      idCol: 1,
      triggerCol: 2,
      prefix: "P-"
    },
    "ApartmentServiceCalc": {
      idCol: 1,
      triggerCol: 2,
      prefix: "ASC-"
    },
    "ApartmentQuoteSummary": {
      idCol: 1,
      triggerCol: 2,
      prefix: "ASUM-"
    }
  };

  if (!config[sheetName]) return;

  const { idCol, triggerCol, prefix } = config[sheetName];

  if (col !== triggerCol) return;

  const triggerValue = sheet.getRange(row, triggerCol).getValue();
  if (!triggerValue) return;

  const idCell = sheet.getRange(row, idCol);
  if (idCell.getValue()) return;

  idCell.setValue(prefix + new Date().getTime());
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (data.action === "sendHoaIntroEmail") {
      const result = sendHoaIntroEmailByLeadId(data.leadId);
      return HtmlService.createHtmlOutput(JSON.stringify(result));
    }

    const quoteId = data.quoteId;

    if (!quoteId) {
      throw new Error("Missing quoteId.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const quotesSheet = ss.getSheetByName("Quotes");
    const lineItemsSheet = ss.getSheetByName("QuoteLineItems");
    const apartmentSummarySheet = ss.getSheetByName("ApartmentQuoteSummary");
    const apartmentServiceCalcSheet = ss.getSheetByName("ApartmentServiceCalc");

    if (!quotesSheet) throw new Error('Sheet not found: "Quotes"');
    if (!lineItemsSheet) throw new Error('Sheet not found: "QuoteLineItems"');
    if (!apartmentSummarySheet) throw new Error('Sheet not found: "ApartmentQuoteSummary"');
    if (!apartmentServiceCalcSheet) throw new Error('Sheet not found: "ApartmentServiceCalc"');

    const quotesData = quotesSheet.getDataRange().getValues();
    const lineItemsData = lineItemsSheet.getDataRange().getValues();
    const apartmentSummaryData = apartmentSummarySheet.getDataRange().getValues();
    const apartmentServiceCalcData = apartmentServiceCalcSheet.getDataRange().getValues();

    if (quotesData.length < 2) throw new Error("Quotes sheet has no data.");

    const quoteHeaders = quotesData[0];
    const lineItemHeaders = lineItemsData[0] || [];
    const apartmentSummaryHeaders = apartmentSummaryData[0] || [];
    const apartmentServiceCalcHeaders = apartmentServiceCalcData[0] || [];

    const quoteRowIndex = findRowIndexById(quotesData, quoteId);
    if (quoteRowIndex === -1) {
      throw new Error("Quote not found for QuoteID: " + quoteId);
    }

    const quote = rowToObject(quoteHeaders, quotesData[quoteRowIndex]);

    if (!quote.Email || !quote.ContactName || !quote.PropertyName) {
      throw new Error("Quote is missing Email, ContactName, or PropertyName.");
    }

    let lineItems = [];
    let effectiveBundleTotal = quote.BundleTotal || "";
    let itemizedTotalValue = quote.Subtotal || "";
    let investmentValue = quote.BundleTotal || "";

    if (String(quote.PropertyType || "").trim() === "Apartment") {
      lineItems = apartmentServiceCalcData
        .slice(1)
        .map(row => rowToObject(apartmentServiceCalcHeaders, row))
        .filter(item => String(item.QuoteID) === String(quoteId))
        .map(mapApartmentCalcToLineItem);

      if (lineItems.length === 0) {
        throw new Error("No apartment service calc rows found for QuoteID: " + quoteId);
      }

      effectiveBundleTotal = getApartmentBundleTotal(
        apartmentSummaryData,
        apartmentSummaryHeaders,
        quoteId
      );

      itemizedTotalValue = getApartmentMaximumPriceTotal(
        apartmentSummaryData,
        apartmentSummaryHeaders,
        quoteId
      );

      investmentValue = getApartmentInvestmentValue(
        apartmentSummaryData,
        apartmentSummaryHeaders,
        quoteId
      );
    } else {
      lineItems = lineItemsData
        .slice(1)
        .map(row => rowToObject(lineItemHeaders, row))
        .filter(item => String(item.QuoteID) === String(quoteId))
        .map(mapQuoteLineItemToLineItem);

      if (lineItems.length === 0) {
        throw new Error("No line items found for QuoteID: " + quoteId);
      }

      itemizedTotalValue = quote.Subtotal || "";
      investmentValue = quote.BundleTotal || quote.Subtotal || "";
    }

    const quoteForPdf = {
      ...quote,
      BundleTotal: effectiveBundleTotal,
      ItemizedTotal: itemizedTotalValue,
      Investment: investmentValue
    };

    const proposalTemplateId = getProposalTemplateIdByPropertyType(quote.PropertyType);
    const quoteTemplateId = getQuoteTemplateIdByPropertyType(quote.PropertyType);

    const proposalPdf = buildPdfFromTemplate(
      proposalTemplateId,
      "Proposal",
      quoteForPdf,
      lineItems
    );

    const quotePdf = buildPdfFromTemplate(
      quoteTemplateId,
      "Quote",
      quoteForPdf,
      lineItems
    );

    setCellByHeaderIfExists(
      quotesSheet,
      quoteHeaders,
      quoteRowIndex + 1,
      "PDFFile",
      proposalPdf.getUrl()
    );

    setCellByHeaderIfExists(
      quotesSheet,
      quoteHeaders,
      quoteRowIndex + 1,
      "ProposalPDFFile",
      proposalPdf.getUrl()
    );

    setCellByHeaderIfExists(
      quotesSheet,
      quoteHeaders,
      quoteRowIndex + 1,
      "QuotePDFFile",
      quotePdf.getUrl()
    );

    const emailContent = getEmailContentByPropertyType(quote);
    const w9CoiFile = DriveApp.getFileById(W9_COI_FILE_ID);

    GmailApp.sendEmail(quote.Email, emailContent.subject, emailContent.body, {
      attachments: [
        proposalPdf.getBlob(),
        quotePdf.getBlob(),
        w9CoiFile.getBlob()
      ],
      name: "Gridiron Pressure Washing"
    });

    setCellByHeader(
      quotesSheet,
      quoteHeaders,
      quoteRowIndex + 1,
      "QuoteStatus",
      "Sent"
    );

    setCellByHeader(
      quotesSheet,
      quoteHeaders,
      quoteRowIndex + 1,
      "SentDate",
      new Date()
    );

    return HtmlService.createHtmlOutput(
      JSON.stringify({
        status: "ok",
        message: "Quote, proposal, and W9/COI attachment sent.",
        quoteId: quoteId,
        proposalPdfUrl: proposalPdf.getUrl(),
        quotePdfUrl: quotePdf.getUrl()
      })
    );
  } catch (err) {
    return HtmlService.createHtmlOutput(
      JSON.stringify({
        status: "error",
        message: err.message
      })
    );
  }
}

function sendHoaIntroEmailByLeadId(leadId) {
  if (!leadId) {
    throw new Error("Missing leadId.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadSheet = ss.getSheetByName("Leadlist");

  if (!leadSheet) {
    throw new Error('Sheet not found: "Leadlist"');
  }

  const data = leadSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("Leadlist sheet has no data.");
  }

  const headers = data[0];
  const rowIndex = findRowIndexById(data, leadId);

  if (rowIndex === -1) {
    throw new Error("Lead not found for LeadID: " + leadId);
  }

  const lead = rowToObject(headers, data[rowIndex]);

  if (!lead.Email) {
    throw new Error("Lead is missing Email.");
  }

  const contactName = safeValue(lead.ContactName) || "there";
  const hoaName = safeValue(lead.BusinessName || lead.PropertyName || "your HOA");

  const subject = "Exterior Cleaning Services for " + hoaName;

  const body =
    "Hi " + contactName + ",\n\n" +
    "I hope you’re doing well. My name is Justin with Gridiron Pressure Washing. I wanted to reach out and introduce our company as a local exterior cleaning option for " + hoaName + ".\n\n" +
    "We work with communities to help maintain a cleaner overall appearance across both shared neighborhood areas and visible exterior surfaces. Depending on the needs of the HOA, that can include house soft washing, sidewalks, curbs, entry signage, mail areas, retaining walls, amenity spaces, and other community-maintained surfaces.\n\n" +
    "With spring conditions and seasonal buildup, keeping the neighborhood clean can make a big difference in curb appeal and overall presentation. Our focus is reliable service, professional communication, and competitive pricing.\n\n" +
    "If it would be helpful, I’d be happy to send over more information or put together a quote for " + hoaName + " for any current or future exterior cleaning needs.\n\n" +
    "Please let me know if you’re the right contact for this, or if there’s someone else I should reach out to.\n\n" +
    "Thank you,\n" +
    "Jaden Gaines & Justin Jefferson\n" +
    "Gridiron Pressure Washing LLC\n" +
    "(470) 632-5817\n" +
    "gridironpressurewash@gmail.com";

  GmailApp.sendEmail(lead.Email, subject, body, {
    name: "Gridiron Pressure Washing"
  });

  setCellByHeaderIfExists(leadSheet, headers, rowIndex + 1, "IntroEmailSent", true);
  setCellByHeaderIfExists(leadSheet, headers, rowIndex + 1, "IntroEmailSentDate", new Date());

  return {
    status: "ok",
    message: "HOA intro email sent.",
    leadId: leadId
  };
}

function buildPdfFromTemplate(templateId, docType, quote, lineItems) {
  const templateFile = DriveApp.getFileById(templateId);
  const docCopy = templateFile.makeCopy(
    docType + " - " + quote.PropertyName + " - " + quote.QuoteID
  );

  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  body.replaceText("{{QuoteID}}", safeValue(quote.QuoteID));
  body.replaceText("{{LeadRef}}", safeValue(quote.LeadRef));
  body.replaceText("{{PropertyType}}", safeValue(quote.PropertyType));
  body.replaceText("{{PropertyName}}", safeValue(quote.PropertyName));
  body.replaceText("{{ContactName}}", safeValue(quote.ContactName));
  body.replaceText("{{Phone}}", safeValue(quote.Phone));
  body.replaceText("{{Email}}", safeValue(quote.Email));
  body.replaceText("{{Address}}", safeValue(quote.Address));
  body.replaceText("{{QuoteDate}}", formatDateValue(quote.QuoteDate));
  body.replaceText("{{TemplateType}}", safeValue(quote.PropertyType));
  body.replaceText("{{DeliveryMethod}}", safeValue(quote.DeliveryMethod));
  body.replaceText("{{Subtotal}}", safeMoney(quote.Subtotal));
  body.replaceText("{{BundleTotal}}", safeMoney(quote.BundleTotal));
  body.replaceText("{{ITEMIZED_TOTAL}}", safeMoney(quote.ItemizedTotal));
  body.replaceText("{{INVESTMENT}}", safeMoney(quote.Investment));
  body.replaceText("{{QuoteStatus}}", safeValue(quote.QuoteStatus));
  body.replaceText("{{Notes}}", safeValue(quote.Notes));

  const scopeOfWorkText = buildScopeOfWorkText(lineItems);
  body.replaceText("{{SCOPE_OF_WORK}}", scopeOfWorkText);

  insertLineItemsTable(body, lineItems);

  doc.saveAndClose();

  const pdfBlob = docCopy.getAs(MimeType.PDF);
  const pdfFile = DriveApp.createFile(pdfBlob).setName(docCopy.getName() + ".pdf");

  docCopy.setTrashed(true);

  return pdfFile;
}

function insertLineItemsTable(body, lineItems) {
  const marker = body.findText("{{LINE_ITEMS}}");
  if (!marker) return;

  const markerElement = marker.getElement();
  const markerText = markerElement.asText();
  markerText.setText("");

  const markerParagraph = markerElement.getParent().asParagraph();
  const markerIndex = body.getChildIndex(markerParagraph);

  const tableData = [
    ["Service", "Scope", "Price", "Notes"]
  ];

  lineItems.forEach(item => {
    tableData.push([
      safeValue(item.ServiceType),
      safeValue(item.Scope),
      safeMoney(item.Price),
      safeValue(item.Notes)
    ]);
  });

  body.insertTable(markerIndex + 1, tableData);
}

function buildScopeOfWorkText(lineItems) {
  const serviceNames = lineItems
    .map(item => String(item.ServiceType || "").trim())
    .filter(name => name);

  const uniqueServices = [...new Set(serviceNames)];

  if (uniqueServices.length === 0) {
    return "Exterior cleaning services as outlined below.";
  }

  const formattedServices = uniqueServices.map(formatServiceNameForScope);
  return formattedServices.map(service => "• " + service).join("\n");
}

function formatServiceNameForScope(serviceName) {
  const raw = String(serviceName || "").trim();
  const cleaned = raw.toLowerCase();

  const map = {
    "soft wash": "Building soft wash",
    "building soft wash": "Building soft wash",
    "building soft wash cleaning": "Building soft wash",
    "breezeways": "Breezeway cleaning",
    "breezeway": "Breezeway cleaning",
    "breezeway cleaning": "Breezeway cleaning",
    "stairs": "Stairwell cleaning",
    "stairwells": "Stairwell cleaning",
    "stairwell cleaning": "Stairwell cleaning",
    "sidewalks": "Sidewalk and walkway cleaning",
    "walkways": "Sidewalk and walkway cleaning",
    "sidewalks/walkways": "Sidewalk and walkway cleaning",
    "sidewalk and walkway cleaning": "Sidewalk and walkway cleaning",
    "curbs": "Curb cleaning",
    "curb cleaning": "Curb cleaning",
    "retaining walls": "Retaining wall cleaning",
    "retaining wall": "Retaining wall cleaning",
    "retaining wall cleaning": "Retaining wall cleaning",
    "common areas": "Common area cleaning",
    "common area": "Common area cleaning",
    "common area cleaning": "Common area cleaning",
    "mail center": "Mail center cleaning",
    "mail center cleaning": "Mail center cleaning",
    "other services": "Other services",
    "patios": "Patio cleaning",
    "patio": "Patio cleaning",
    "patio cleaning": "Patio cleaning"
  };

  if (map[cleaned]) return map[cleaned];

  if (cleaned.endsWith(" cleaning")) {
    return toTitleCase(raw);
  }

  return toTitleCase(raw) + " cleaning";
}

function formatServiceNameForQuote(serviceName) {
  const cleaned = String(serviceName || "").trim().toLowerCase();

  const map = {
    "soft wash": "Building Soft Wash",
    "building soft wash": "Building Soft Wash",
    "breezeways": "Breezeway Cleaning",
    "breezeway": "Breezeway Cleaning",
    "stairs": "Stairwell Cleaning",
    "stairwells": "Stairwell Cleaning",
    "sidewalks": "Sidewalk and Walkway Cleaning",
    "walkways": "Sidewalk and Walkway Cleaning",
    "sidewalks/walkways": "Sidewalk and Walkway Cleaning",
    "curbs": "Curb Cleaning",
    "retaining walls": "Retaining Wall Cleaning",
    "retaining wall": "Retaining Wall Cleaning",
    "common areas": "Common Area Cleaning",
    "common area": "Common Area Cleaning",
    "mail center": "Mail Center Cleaning",
    "mail room": "Mail Room Cleaning",
    "other services": "Other Services",
    "patios": "Patio Cleaning",
    "patio": "Patio Cleaning"
  };

  return map[cleaned] || toTitleCase(serviceName);
}

function buildServiceScopeDescription(serviceType, quantity) {
  const cleaned = String(serviceType || "").trim().toLowerCase();

  const map = {
    "soft wash": "Soft washing of building exterior surfaces to remove visible organic buildup and improve the overall appearance of the property.",
    "building soft wash": "Soft washing of building exterior surfaces to remove visible organic buildup and improve the overall appearance of the property.",
    "breezeways": "Cleaning of breezeway surfaces to remove dirt, buildup, and staining from resident access areas.",
    "breezeway": "Cleaning of breezeway surfaces to remove dirt, buildup, and staining from resident access areas.",
    "stairs": "Cleaning of stairwells and landings to improve cleanliness, appearance, and the presentation of high-traffic access points.",
    "stairwells": "Cleaning of stairwells and landings to improve cleanliness, appearance, and the presentation of high-traffic access points.",
    "sidewalks": "Surface cleaning of sidewalks and walkways to remove buildup and improve the appearance of common paths throughout the property.",
    "walkways": "Surface cleaning of sidewalks and walkways to remove buildup and improve the appearance of common paths throughout the property.",
    "sidewalks/walkways": "Surface cleaning of sidewalks and walkways to remove buildup and improve the appearance of common paths throughout the property.",
    "curbs": "Cleaning of curb lines to sharpen the look of traffic areas and help the property appear cleaner and better maintained.",
    "other services": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "retaining walls": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "retaining wall": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "common areas": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "common area": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "mail center": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces.",
    "mail room": "Additional exterior cleaning services for supporting areas such as mail rooms, retaining walls, common areas, and similar shared property surfaces."
  };

  return map[cleaned] || "Exterior cleaning service as outlined in the proposed scope of work.";
}

function buildGenericScopeDescription(item) {
  const qty = item.Quantity || "";
  const unitType = item.UnitType || "";

  if (qty && unitType) return qty + " " + unitType;
  if (qty) return String(qty);
  return "Service as outlined.";
}

function toTitleCase(text) {
  return String(text || "")
    .toLowerCase()
    .split(" ")
    .map(word => word ? word.charAt(0).toUpperCase() + word.slice(1) : "")
    .join(" ");
}

function getProposalTemplateIdByPropertyType(propertyType) {
  const map = {
    "Apartment": "1s9KhAm7ii35Z16oiHhmIag-mbQH3QS3OTcfLYkWS_WY",
    "Residential": "1wnUZIrwkIymYq3SfQjbUkTh-2kbZfuXYqdqCGrSW7Kk",
    "HOA": "1dj_sLE8XDjRPZOaBD_-XI4Gd6RV9HXkqoyumcW6CrgE",
    "Commercial": "1jzOXkXqbZrVeIjDSJ9CUX8egSx4WLTTP_zE1UzyAqW8"
  };

  const templateId = map[String(propertyType || "").trim()];
  if (!templateId) {
    throw new Error("No proposal template mapped for PropertyType: " + propertyType);
  }

  return templateId;
}

function getQuoteTemplateIdByPropertyType(propertyType) {
  const map = {
    "Apartment": "1I_up-Ql2Kvqxr8r03BNWVu57Gu6inCh4-3O60OEUvlg",
    "Residential": "1Pf8gCbdfSLTMUGjSkYs6gKlfHKjTx6aeFcrFZ_W0rSo",
    "HOA": "12ozpIppac4zuWdfyS49Cq1rrWGUPlCaDqbsynY_UWfM",
    "Commercial": "1A92CfsJ7pQrbaHq2mMV7t24rIN2MvF_W_37QdEAeYb8"
  };

  const templateId = map[String(propertyType || "").trim()];
  if (!templateId) {
    throw new Error("No quote template mapped for PropertyType: " + propertyType);
  }

  return templateId;
}

function getEmailContentByPropertyType(quote) {
  const propertyType = String(quote.PropertyType || "").trim();
  const contactName = safeValue(quote.ContactName);
  const propertyName = safeValue(quote.PropertyName);

  if (propertyType === "Residential") {
    return {
      subject: "Exterior Cleaning Quote for " + propertyName,
      body:
        "Hi " + contactName + ",\n\n" +
        "Thank you for taking the time to speak with me.\n\n" +
        "Attached is your quote for the exterior cleaning services we discussed. If everything looks good to you, let me know and I can get you on the schedule.\n\n" +
        "If you have any questions, I’d be happy to help.\n\n" +
        "Thank you,\n\n" +
        "Jaden Gaines & Justin Jefferson\n" +
        "Gridiron Pressure Washing\n" +
        "(470) 632-5817\n" +
        "gridironpressurewash@gmail.com"
    };
  }

  if (propertyType === "Apartment") {
    return {
      subject: "Exterior Cleaning Proposal for " + propertyName,
      body:
        "Hi " + contactName + ",\n\n" +
        "I appreciate you taking the time to speak with me.\n\n" +
        "Attached is the exterior cleaning proposal for " + propertyName + ". I put this together so you have a clear breakdown of the scope and pricing for the services discussed.\n\n" +
        "A lot of communities are simply looking for the same or better value while keeping the property looking its best. Our process is designed to help improve curb appeal, remove buildup that water alone typically does not fully address, and save your team time by allowing us to handle the exterior cleaning efficiently.\n\n" +
        "If there is anything you would like adjusted, I would be happy to tailor the proposal around your priorities or budget. Even if you already have a vendor or currently handle it in-house, this can still give you another option for comparison.\n\n" +
        "Let me know if you have any questions.\n\n" +
        "Thank you,\n\n" +
        "Jaden Gaines & Justin Jefferson\n" +
        "Gridiron Pressure Washing\n" +
        "(470) 632-5817\n" +
        "gridironpressurewash@gmail.com"
    };
  }

  if (propertyType === "HOA") {
    return {
      subject: "Community Exterior Cleaning Proposal for " + propertyName,
      body:
        "Hi " + contactName + ",\n\n" +
        "I appreciate you taking the time to speak with me.\n\n" +
        "Attached is the exterior cleaning proposal for " + propertyName + ". I put this together so you have a clear outline of the recommended scope and pricing for the services discussed.\n\n" +
        "We understand how important it is for an HOA to maintain a clean and well-presented community. Our process is designed to help improve the appearance of common areas, support overall curb appeal, and provide a dependable solution for exterior cleaning needs.\n\n" +
        "If you would like any part of the proposal adjusted to better fit the community’s priorities or budget, I would be happy to revise it.\n\n" +
        "Please feel free to reach out with any questions. I would be glad to go over the proposal with you and discuss next steps.\n\n" +
        "Thank you,\n\n" +
        "Jaden Gaines & Justin Jefferson\n" +
        "Gridiron Pressure Washing\n" +
        "(470) 632-5817\n" +
        "gridironpressurewash@gmail.com"
    };
  }

  if (propertyType === "Commercial") {
    return {
      subject: "Commercial Exterior Cleaning Proposal for " + propertyName,
      body:
        "Hi " + contactName + ",\n\n" +
        "I appreciate you taking the time to speak with me.\n\n" +
        "Attached is the exterior cleaning proposal for " + propertyName + ". I put this together so you have a clear outline of the scope and pricing for the services discussed.\n\n" +
        "We understand how important it is for a commercial property to maintain a clean and professional appearance for customers, tenants, and visitors. Our process is designed to help improve the overall appearance of the property, remove buildup from exterior surfaces, and provide a dependable solution for ongoing exterior cleaning needs.\n\n" +
        "If you would like anything adjusted to better fit your priorities, schedule, or budget, I would be happy to revise the proposal.\n\n" +
        "Please feel free to reach out with any questions. I would be glad to go over the proposal with you and discuss next steps.\n\n" +
        "Thank you,\n\n" +
        "Jaden Gaines & Justin Jefferson\n" +
        "Gridiron Pressure Washing\n" +
        "(470) 632-5817\n" +
        "gridironpressurewash@gmail.com"
    };
  }

  throw new Error("No email content mapped for PropertyType: " + propertyType);
}

function getApartmentBundleTotal(apartmentSummaryData, apartmentSummaryHeaders, quoteId) {
  for (let i = 1; i < apartmentSummaryData.length; i++) {
    const rowObj = rowToObject(apartmentSummaryHeaders, apartmentSummaryData[i]);

    if (String(rowObj.QuoteID) === String(quoteId)) {
      const bundlePrice = Number(rowObj.BundlePrice || 0);
      const totalCost = Number(rowObj.TotalCost || 0);
      const maximumPriceTotal = Number(rowObj.MaximumPriceTotal || 0);

      if (bundlePrice < totalCost) {
        return maximumPriceTotal || "";
      }

      return bundlePrice || maximumPriceTotal || "";
    }
  }
  return "";
}

function getApartmentMaximumPriceTotal(apartmentSummaryData, apartmentSummaryHeaders, quoteId) {
  for (let i = 1; i < apartmentSummaryData.length; i++) {
    const rowObj = rowToObject(apartmentSummaryHeaders, apartmentSummaryData[i]);
    if (String(rowObj.QuoteID) === String(quoteId)) {
      return rowObj.MaximumPriceTotal || "";
    }
  }
  return "";
}

function getApartmentInvestmentValue(apartmentSummaryData, apartmentSummaryHeaders, quoteId) {
  for (let i = 1; i < apartmentSummaryData.length; i++) {
    const rowObj = rowToObject(apartmentSummaryHeaders, apartmentSummaryData[i]);

    if (String(rowObj.QuoteID) === String(quoteId)) {
      const bundlePrice = Number(rowObj.BundlePrice || 0);
      const totalCost = Number(rowObj.TotalCost || 0);
      const maximumPriceTotal = Number(rowObj.MaximumPriceTotal || 0);

      if (bundlePrice < totalCost) {
        return maximumPriceTotal || "";
      }

      return bundlePrice || maximumPriceTotal || "";
    }
  }
  return "";
}

function mapQuoteLineItemToLineItem(item) {
  return {
    ServiceType: item.ServiceType || "",
    Scope: buildGenericScopeDescription(item),
    Price: item.LineTotal || item.UnitPrice || "",
    Notes: item.Notes || ""
  };
}

function mapApartmentCalcToLineItem(item) {
  const serviceType = item.ServiceType || "";

  return {
    ServiceType: formatServiceNameForQuote(serviceType),
    Scope: buildServiceScopeDescription(serviceType, item.SquareFootage),
    Price: item.QuotedPrice || "",
    Notes: item.Notes || ""
  };
}

function findRowIndexById(data, quoteId) {
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(quoteId)) {
      return i;
    }
  }
  return -1;
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((header, i) => {
    obj[header] = row[i];
  });
  return obj;
}

function setCellByHeader(sheet, headers, rowNumber, headerName, value) {
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) {
    throw new Error("Column not found: " + headerName);
  }
  sheet.getRange(rowNumber, colIndex + 1).setValue(value);
}

function setCellByHeaderIfExists(sheet, headers, rowNumber, headerName, value) {
  const colIndex = headers.indexOf(headerName);
  if (colIndex === -1) return;
  sheet.getRange(rowNumber, colIndex + 1).setValue(value);
}

function safeValue(value) {
  return value === null || value === undefined ? "" : String(value);
}

function safeMoney(value) {
  if (value === null || value === undefined || value === "") return "";
  const num = Number(value);
  if (!isNaN(num)) {
    return "$" + num.toFixed(2);
  }
  return "$" + String(value);
}

function formatDateValue(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "M/d/yyyy");
  }
  return String(value);
}

function testAuth() {
  GmailApp.getInboxThreads(0, 1);
  DriveApp.getRootFolder();

  const doc = DocumentApp.create("Auth Test Temp");
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  SpreadsheetApp.getActiveSpreadsheet();
}
