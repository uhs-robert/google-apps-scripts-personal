// sheets/time-tracker/email-invoice.js
/**
 * Sends a professional HTML invoice email with the PDF attached.
 * @param {Object} invoiceData - Invoice data from getInvoiceByRow()
 * @param {string} invoiceNumber - Generated invoice number
 * @param {boolean} includeFee - Whether processing fee was included
 * @param {Blob} pdfBlob - PDF blob to attach
 * @param {string} pdfUrl - Drive URL for the "View Invoice PDF" button
 */
function sendInvoiceEmail(
  invoiceData,
  invoiceNumber,
  includeFee,
  pdfBlob,
  pdfUrl,
) {
  const processingFee = includeFee
    ? invoiceData.gross * PROCESSING_FEE_PERCENTAGE
    : 0;
  const totalDue = invoiceData.gross + processingFee;

  const fmt = (v) => {
    const n = typeof v === "string" ? parseFloat(v) : v;
    return (isNaN(n) ? 0 : n).toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  };

  const tz = Session.getScriptTimeZone();
  const invoiceDate = new Date();
  const dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + DUE_DATE_DAYS);

  const template = HtmlService.createTemplateFromFile("InvoiceEmail");
  template.invoiceNumber = invoiceNumber;
  template.client = invoiceData.client;
  template.projectTitle = invoiceData.projectTitle;
  template.invoiceDateStr = Utilities.formatDate(
    invoiceDate,
    tz,
    "MMMM dd, yyyy",
  );
  template.dueDateStr = Utilities.formatDate(dueDate, tz, "MMMM dd, yyyy");
  template.subtotal = fmt(invoiceData.gross);
  template.processingFee = fmt(processingFee);
  template.totalDue = fmt(totalDue);
  template.includeFee = includeFee;
  template.lineItemTitle = LINE_ITEM_TITLE;
  template.hoursUsed = invoiceData.hoursUsed;
  template.rate = fmt(invoiceData.rate);
  template.paymentTerms = PAYMENT_TERMS;
  template.pdfUrl = pdfUrl;

  const htmlBody = template.evaluate().getContent();
  const subject = `Invoice ${invoiceNumber} from UpHill Solutions`;

  const logoBlob = DriveApp.getFileById('1ja5S94DUdx_lj5tbbUBepNMhnK39IW2B').getBlob();

  const recipient = invoiceData.emailTo || Session.getActiveUser().getEmail();
  GmailApp.sendEmail(
    recipient,
    subject,
    `Please find your invoice ${invoiceNumber} attached. You can also view it online: ${pdfUrl}`,
    {
      htmlBody: htmlBody,
      attachments: [
        pdfBlob.setName(`${invoiceNumber} - ${invoiceData.client}.pdf`),
      ],
      inlineImages: { logo: logoBlob },
      name: "UpHill Solutions",
      replyTo: Session.getActiveUser().getEmail(),
    },
  );

  Logger.log(
    `Invoice email sent to ${recipient} — subject: ${subject}`,
  );
}
