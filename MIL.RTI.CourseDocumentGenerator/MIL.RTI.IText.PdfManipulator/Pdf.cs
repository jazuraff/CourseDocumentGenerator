using System.Collections.Generic;
using iText.Forms;
using iText.Kernel.Pdf;

namespace MIL.RTI.IText.PdfManipulator
{
    public abstract class Pdf
    {
        private readonly string _pdfDestination;
        private readonly string _pdfSource;

        protected Pdf(string pdfSource, string pdfDestination)
        {
            _pdfSource = pdfSource;
            _pdfDestination = pdfDestination;
        }

        private PdfDocument GetPdf()
        {
            var reader = new PdfReader(_pdfSource).SetUnethicalReading(true);
            var writer = new PdfWriter(_pdfDestination);
            var document = new PdfDocument(reader, writer);

            return document;
        }

        public void ManipulateFields(Dictionary<string, string> fields)
        {
            var document = GetPdf();
            var form = PdfAcroForm.GetAcroForm(document, true);

            foreach (var field in fields) form.GetField(field.Key).SetValue(field.Value);

            form.RemoveXfaForm();
            document.GetCatalog().Remove(PdfName.Perms);

            document.Close();
        }

        /**
         * Removes any usage rights that this PDF may have. Only Adobe can grant usage rights
         * and any PDF modification with iText will invalidate them. Invalidated usage rights may
         * confuse Acrobat and it's advisable to remove them altogether.
         */
        protected void RemoveUsageRights(PdfDocument document)
        {
            var perms = document.GetCatalog().GetPdfObject().GetAsDictionary(PdfName.Perms);

            if (perms == null) return;
            perms.Remove(new PdfName("UR"));
            perms.Remove(PdfName.UR3);
            if (perms.Size() == 0) document.GetCatalog().Remove(PdfName.Perms);
        }
    }
}